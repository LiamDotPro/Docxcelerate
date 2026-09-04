# Probe C -- Word's own reading of a conformance case.
#
# Drives Word 16 over COM, walks the body's paragraphs, and emits ONE line of
# compact JSON on stdout. Word is the only engine whose numbers count as ground
# truth here: the file and the preview can both be "right" while Word lays the
# page out differently, and it is Word the document is for.
#
# Every measurement sits in its own try/catch. A property Word will not give up
# records null, never a throw -- a probe that dies on one paragraph costs the
# whole board its data, and "not available" is a measurement.
#
# PowerShell 5.1: no && chains, no ternary. The body is ASCII-only so the file
# parses identically whether 5.1 reads it as ANSI or UTF-8.

param(
    [Parameter(Mandatory = $true)][string]$DocxPath,
    [Parameter(Mandatory = $true)][string]$PdfPath
)

$ErrorActionPreference = 'Stop'

# PS 5.1 writes redirected stdout in the console codepage, which mangles any
# non-ASCII the document holds. Emit UTF-8 so the JSON survives the pipe.
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

# Word constants.
$wdInfoPage = 3           # wdActiveEndPageNumber
$wdInfoX = 5              # horizontal position relative to the page, in points
$wdInfoY = 6              # vertical position relative to the page, in points
$wdStatPages = 2
$wdCollapseStart = 1
$wdCollapseEnd = 0
$wdExportPdf = 17
$wdDoNotSaveChanges = 0
$wdPrintView = 3
$wdSeekMainDocument = 0
$wdSeekCurrentPageHeader = 9
$wdSeekCurrentPageFooter = 10
$wdLineSpaceSingle = 0
$wdLineSpaceAtLeast = 3
$wdLineSpaceExactly = 4
$wdLineSpaceMultiple = 5
$wdUndefined = 9999999

# Alignment, as Word numbers it.
$alignNames = @{ 0 = 'left'; 1 = 'center'; 2 = 'right'; 3 = 'justify' }
# How a cell's content sits against the height of the cell.
$vAlignNames = @{ 0 = 'top'; 1 = 'center'; 3 = 'bottom' }
# How a row's height is meant: taken from the content, a floor, or a ceiling.
$rowHeightRules = @{ 0 = 'auto'; 1 = 'atLeast'; 2 = 'exactly' }
# Where a table sits across the text column.
$rowAlignNames = @{ 0 = 'left'; 1 = 'center'; 2 = 'right' }
# Tab alignment and leader, likewise.
$tabAlignNames = @{ 0 = 'left'; 1 = 'center'; 2 = 'right'; 3 = 'decimal'; 4 = 'bar' }
$tabLeaderNames = @{ 0 = 'spaces'; 1 = 'dot'; 2 = 'dash'; 3 = 'line'; 4 = 'heavy'; 5 = 'middleDot' }

# Where a range actually sits on the page.
#
# It has to go through Selection, not through Range.Information, and not as a
# fallback -- as the only path. Measured on a page of five paragraphs set five
# ways, Range.Information(5) returned the left margin for every one of them,
# including the centred and the right-ranged: it reports where the *line* may
# begin, not where the glyphs do. Selecting the range first and asking the
# Selection returns 57pt / 248pt / 419pt for left, centred and right, which is
# what the page shows.
#
# Range.Information also returns -1 for a range Word has not laid out yet, and
# selecting is what forces the layout. So the same call fixes both, and a value
# still negative afterwards is Word saying "cannot determine" -- recorded as
# null, never as a fake coordinate.
function Get-Info {
    param($range, [int]$index)
    $value = $null
    try {
        $null = $range.Select()
        $value = [double]$script:w.Selection.Information($index)
    } catch {
        # A range that will not take a selection is still worth asking about
        # directly; the answer is coarse but it is not nothing.
        try { $value = [double]$range.Information($index) } catch { return $null }
    }
    if ($null -eq $value -or $value -lt 0) { return $null }
    return [math]::Round($value, 2)
}

# A value Word reports as "the setting is mixed or undefined".
function Clean-Undefined {
    param($value)
    if ($null -eq $value) { return $null }
    $number = [double]$value
    if ($number -eq $wdUndefined -or $number -eq -$wdUndefined) { return $null }
    return [math]::Round($number, 2)
}

# Word's tri-state booleans: -1 true, 0 false, 9999999 undefined.
function Clean-Bool {
    param($value)
    if ($null -eq $value) { return $null }
    $number = [int]$value
    if ($number -eq $wdUndefined) { return $null }
    return ($number -ne 0)
}

# A colour as the six hex digits the rest of the harness speaks. Word hands back
# a BGR integer, and wdColorAutomatic (-16777216) means "no colour was set".
function To-Hex {
    param($value)
    if ($null -eq $value) { return $null }
    $number = [int64]$value
    if ($number -lt 0) { return $null }
    $b = ($number -shr 16) -band 255
    $g = ($number -shr 8) -band 255
    $r = $number -band 255
    return ('{0:X2}{1:X2}{2:X2}' -f $r, $g, $b)
}

# Paragraph text without Word's trailing carriage return and cell markers.
function Get-ParaText {
    param($paragraph)
    $text = $paragraph.Range.Text
    if ($null -eq $text) { return '' }
    return $text.TrimEnd([char]13, [char]7)
}


# One table, and everything Word will say about it.
#
# The cells are walked through the table's RANGE, not through Rows and Columns.
# Table.Cell(r, c), Table.Columns and Table.Rows are all documented to fail on
# a table whose cells do not line up -- "cannot access individual rows in this
# collection because the table has mixed cell widths" -- and that is exactly
# the table a merge case is about, so a probe reading cells that way would
# report nothing for the documents it was written for. Range.Cells does not
# depend on the table being a plain grid, and every cell it yields carries the
# RowIndex and ColumnIndex the other reading would have given.
#
# Rows are still walked, in their own try, for the handful of properties that
# belong to a row rather than to a cell. Coming back empty is itself a
# measurement.
function Read-Table {
    param($table, [int]$index, [string]$path, [int]$depth)

    $record = [ordered]@{
        index              = $index
        path               = $path
        depth              = $depth
        rowCount           = $null
        columnCount        = $null
        uniform            = $null
        style              = $null
        preferredWidth     = $null
        preferredWidthType = $null
        alignment          = $null
        leftIndent         = $null
        wrapAroundText     = $null
        allowAutoFit       = $null
        page               = $null
        x                  = $null
        y                  = $null
        rows               = @()
        cells              = @()
        nested             = @()
    }

    try { $record.rowCount = [int]$table.Rows.Count } catch { }
    try { $record.columnCount = [int]$table.Columns.Count } catch { }
    # False for a table whose cells do not all line up -- which is what a
    # merged or spanned table is, and the fact a merge case asks Word for.
    try { $record.uniform = [bool]$table.Uniform } catch { }
    try { $record.style = [string]$table.Style.NameLocal } catch { }
    try { $record.preferredWidth = Clean-Undefined $table.PreferredWidth } catch { }
    try { $record.preferredWidthType = [int]$table.PreferredWidthType } catch { }
    try { $record.alignment = $rowAlignNames[[int]$table.Rows.Alignment] } catch { }
    try { $record.leftIndent = [math]::Round([double]$table.Rows.LeftIndent, 2) } catch { }
    # Whether the text runs around the table rather than above and below it.
    try { $record.wrapAroundText = [bool]$table.Rows.WrapAroundText } catch { }
    try { $record.allowAutoFit = [bool]$table.AllowAutoFit } catch { }

    try {
        $start = $table.Range.Duplicate
        $null = $start.Collapse($wdCollapseStart)
        $record.page = [int](Get-Info $start $wdInfoPage)
        $record.x = Get-Info $start $wdInfoX
        $record.y = Get-Info $start $wdInfoY
    } catch { }

    $rows = @()
    try {
        $rowIndex = 0
        foreach ($r in $table.Rows) {
            $rowIndex += 1
            $row = [ordered]@{ index = ($rowIndex - 1) }
            # HeadingFormat is Word's name for w:tblHeader: the row repeats at
            # the top of every page the table runs onto.
            try { $row.headingFormat = Clean-Bool $r.HeadingFormat } catch { $row.headingFormat = $null }
            try { $row.height = Clean-Undefined $r.Height } catch { $row.height = $null }
            try { $row.heightRule = $rowHeightRules[[int]$r.HeightRule] } catch { $row.heightRule = $null }
            try { $row.allowBreakAcrossPages = Clean-Bool $r.AllowBreakAcrossPages } catch { $row.allowBreakAcrossPages = $null }
            try { $row.cellCount = [int]$r.Cells.Count } catch { $row.cellCount = $null }
            $rows += $row
        }
    } catch { }
    $record.rows = @($rows)

    $cells = @()
    try {
        foreach ($c in $table.Range.Cells) {
            $cell = [ordered]@{ row = $null; column = $null; text = '' }

            try { $cell.row = ([int]$c.RowIndex - 1) } catch { }
            try { $cell.column = ([int]$c.ColumnIndex - 1) } catch { }
            # A cell's range ends in the cell marker Word writes, which is two
            # invisible characters that would otherwise be part of every anchor.
            try { $cell.text = ($c.Range.Text -replace "[`a`v`f`n`r]", '').Trim() } catch { }

            try { $cell.width = Clean-Undefined $c.Width } catch { $cell.width = $null }
            try { $cell.preferredWidth = Clean-Undefined $c.PreferredWidth } catch { $cell.preferredWidth = $null }
            try { $cell.height = Clean-Undefined $c.Height } catch { $cell.height = $null }
            try { $cell.heightRule = $rowHeightRules[[int]$c.HeightRule] } catch { $cell.heightRule = $null }
            try { $cell.vAlign = $vAlignNames[[int]$c.VerticalAlignment] } catch { $cell.vAlign = $null }
            try { $cell.shading = To-Hex $c.Shading.BackgroundPatternColor } catch { $cell.shading = $null }

            # The room inside the cell, which Word calls padding and the file
            # calls w:tcMar. In points, like everything else this probe reports.
            try {
                $cell.padding = [ordered]@{
                    top    = [math]::Round([double]$c.TopPadding, 2)
                    right  = [math]::Round([double]$c.RightPadding, 2)
                    bottom = [math]::Round([double]$c.BottomPadding, 2)
                    left   = [math]::Round([double]$c.LeftPadding, 2)
                }
            } catch { $cell.padding = $null }

            try {
                $cell.borders = [ordered]@{
                    top    = ([int]$c.Borders(-1).LineStyle -ne 0)
                    left   = ([int]$c.Borders(-2).LineStyle -ne 0)
                    bottom = ([int]$c.Borders(-3).LineStyle -ne 0)
                    right  = ([int]$c.Borders(-4).LineStyle -ne 0)
                }
            } catch { $cell.borders = $null }

            # How the cell's own paragraphs are set. A column's alignment
            # reaches the page through them, not through the cell.
            try { $cell.alignment = $alignNames[[int]$c.Range.ParagraphFormat.Alignment] } catch { $cell.alignment = $null }

            try {
                $at = $c.Range.Duplicate
                $null = $at.Collapse($wdCollapseStart)
                $cell.page = [int](Get-Info $at $wdInfoPage)
                $cell.x = Get-Info $at $wdInfoX
                $cell.y = Get-Info $at $wdInfoY
            } catch { }

            $cells += $cell
        }
    } catch { }
    $record.cells = @($cells)

    # The tables inside this one's cells. Depth-limited rather than trusted:
    # a document that nested tables forever would take the probe with it.
    if ($depth -lt 3) {
        $nested = @()
        try {
            $at = 0
            foreach ($inner in $table.Tables) {
                $nested += Read-Table $inner $at "$path.$at" ($depth + 1)
                $at += 1
            }
        } catch { }
        $record.nested = @($nested)
    }

    return $record
}

# The shapes Word found, and what it makes of each.
#
# Shapes and InlineShapes are two collections holding the same kind of thing
# with different anchoring: a shape written with no position is inline, one
# given a position floats. A document's shapes are whatever is in both, so
# both are walked and the record says which collection it came from.
#
# `AutoShapeType` is the fact a shape case is really about -- it is what
# separates a drawn rectangle from a picture, and Word reports it only for a
# shape it built from geometry rather than from pixels.
function Read-Shapes {
    param($doc)

    $found = @()

    foreach ($pair in @(@{ items = $doc.Shapes; anchored = 'floating' }, @{ items = $doc.InlineShapes; anchored = 'inline' })) {
        $items = $pair.items
        $count = 0
        try { $count = [int]$items.Count } catch { continue }

        for ($i = 1; $i -le $count; $i += 1) {
            $s = $null
            try { $s = $items.Item($i) } catch { continue }

            $record = [ordered]@{
                index    = $found.Count
                anchored = $pair.anchored
                at       = $null
                name     = $null
                type     = $null
                width    = $null
                height    = $null
                fill     = $null
                line     = $null
                hasText  = $null
                text     = ''
                page     = $null
                x        = $null
                y        = $null
            }

            # Where the shape is anchored in the document, so the collection
            # can be put back into the order the document declares. Word's
            # Shapes collection is in z-order, which for two shapes written one
            # after the other came back reversed -- and a case asking for "the
            # first shape" means the first one written, as every other probe
            # and the preview's own querySelectorAll do.
            try { $record.at = [int]$s.Anchor.Start } catch {
                try { $record.at = [int]$s.Range.Start } catch { }
            }

            try { $record.name = [string]$s.Name } catch { }
            # msoAutoShape is 1 and msoPicture is 13. Naming them is what lets
            # a case say "a real shape" rather than quote a number at a reader.
            try {
                $type = [int]$s.Type
                $typeNames = @{ 1 = 'autoShape'; 13 = 'picture'; 17 = 'textBox'; 3 = 'chart' }
                $record.type = $typeNames[$type]
                if ($null -eq $record.type) { $record.type = "type$type" }
            } catch { }
            try { $record.width = [math]::Round([double]$s.Width, 2) } catch { }
            try { $record.height = [math]::Round([double]$s.Height, 2) } catch { }
            try { $record.fill = To-Hex $s.Fill.ForeColor.RGB } catch { }
            # A shape Word draws no outline on reports Line.Visible false, and
            # asking its colour then returns whatever was last set rather than
            # nothing -- so the visibility is checked before the colour.
            try {
                if ([int]$s.Line.Visible -ne 0) { $record.line = To-Hex $s.Line.ForeColor.RGB }
            } catch { }

            try {
                $frame = $s.TextFrame
                $record.hasText = [bool]$frame.HasText
                if ($record.hasText) {
                    $record.text = ($frame.TextRange.Text -replace "[`a`v`f`n`r]", ' ').Trim()
                }
            } catch { }

            # An inline shape sits in the text and can be asked where it is; a
            # floating one is positioned against the page and reports Top and
            # Left directly. Both are recorded page-relative, in points.
            try {
                $anchor = $s.Anchor.Duplicate
                $null = $anchor.Collapse($wdCollapseStart)
                $record.page = [int](Get-Info $anchor $wdInfoPage)
            } catch { }
            try { $record.x = [math]::Round([double]$s.Left, 2) } catch { }
            try { $record.y = [math]::Round([double]$s.Top, 2) } catch { }

            $found += $record
        }
    }

    # Document order, and renumbered in it. A shape with no anchor Word would
    # give up sorts to the front rather than throwing the whole list out.
    $ordered = @($found | Sort-Object -Property @{ Expression = { if ($null -eq $_.at) { -1 } else { $_.at } } })
    for ($i = 0; $i -lt $ordered.Count; $i += 1) { $ordered[$i].index = $i }

    return @($ordered)
}
$docxFull = (Resolve-Path -LiteralPath $DocxPath).Path
$pdfFull = [System.IO.Path]::GetFullPath($PdfPath)

$out = [ordered]@{
    probe       = 'C'
    ok          = $false
    docx        = $docxFull
    pdf         = $pdfFull
    wordVersion = $null
    pages       = $null
    pageSetup   = $null
    furniture   = $null
    tables      = @()
    shapes      = @()
    paragraphs  = @()
    pdfExported = $false
    errors      = @()
}
$errors = @()

$w = $null
$d = $null

try {
    $w = New-Object -ComObject Word.Application
    $w.Visible = $false
    $w.DisplayAlerts = 0
    $out.wordVersion = [string]$w.Version

    $missing = [System.Reflection.Missing]::Value
    # Read-only, never added to recent files: the harness measures the
    # document, it does not touch it.
    $d = $w.Documents.Open($docxFull, $false, $true, $false,
        $missing, $missing, $missing, $missing, $missing, $missing, $missing, $false)

    # Information(5/6) is page-relative only in Print Layout. In another view
    # Word reports column-relative numbers that look plausible and are wrong.
    try {
        $view = $d.ActiveWindow.View
        if ([int]$view.Type -ne $wdPrintView) { $view.Type = $wdPrintView }
    } catch { $errors += "view: $($_.Exception.Message)" }

    try { $out.pages = [int]$d.ComputeStatistics($wdStatPages) } catch { $errors += "pages: $($_.Exception.Message)" }

    # The page, so the node side can turn page-relative points into
    # content-relative ones and compare like with like against the preview.
    try {
        $ps = $d.PageSetup
        $out.pageSetup = [ordered]@{
            pageWidth    = [math]::Round([double]$ps.PageWidth, 2)
            pageHeight   = [math]::Round([double]$ps.PageHeight, 2)
            leftMargin   = [math]::Round([double]$ps.LeftMargin, 2)
            rightMargin  = [math]::Round([double]$ps.RightMargin, 2)
            topMargin    = [math]::Round([double]$ps.TopMargin, 2)
            bottomMargin = [math]::Round([double]$ps.BottomMargin, 2)
            # How far the running strips sit from the paper's edge, which is a
            # different distance from the margin and is why a header can print
            # above the text without eating into it.
            headerDistance = [math]::Round([double]$ps.HeaderDistance, 2)
            footerDistance = [math]::Round([double]$ps.FooterDistance, 2)
            differentFirstPageHeaderFooter = ([int]$ps.DifferentFirstPageHeaderFooter -ne 0)
            oddAndEvenPagesHeaderFooter = ([int]$ps.OddAndEvenPagesHeaderFooter -ne 0)
        }
    } catch { $errors += "pageSetup: $($_.Exception.Message)" }

    # The running furniture, by the kind of page it is drawn on.
    #
    # Word keeps a header and a footer per type per section, and reports each
    # whether or not it is used: a document that never turned on a first-page
    # header still has an empty one. `exists` is what separates "there is a
    # strip here" from "there is a slot for one", and it is the fact a case
    # about first-page furniture is actually asking about.
    try {
        $furniture = [ordered]@{}
        $section = $d.Sections.Item(1)
        $kinds = @{ primary = 1; firstPage = 2; evenPages = 3 }

        foreach ($kind in $kinds.Keys) {
            $index = $kinds[$kind]

            foreach ($part in @('header', 'footer')) {
                $strip = $null
                try {
                    if ($part -eq 'header') { $strip = $section.Headers.Item($index) }
                    else { $strip = $section.Footers.Item($index) }
                } catch { continue }

                $record = [ordered]@{ exists = $null; text = $null; lines = $null; y = $null }

                try { $record.exists = [bool]$strip.Exists } catch { }
                try {
                    $text = $strip.Range.Text
                    if ($null -ne $text) {
                        $record.text = ($text -replace "[`a`v`f`n`r]", ' ').Trim()
                    }
                } catch { }
                try { $record.lines = [int]$strip.Range.ComputeStatistics(1) } catch { }
                $furniture["$kind.$part"] = $record
            }
        }

        $out.furniture = $furniture
    } catch { $errors += "furniture: $($_.Exception.Message)" }

    # The body's tables. Only the top-level ones are walked here; a table
    # inside a cell is reached through the cell that holds it, which is where
    # it belongs and where the nesting is a fact rather than a flat list.
    try {
        $tables = @()
        $tableIndex = 0
        foreach ($t in $d.Tables) {
            $tables += Read-Table $t $tableIndex "$tableIndex" 0
            $tableIndex += 1
        }
        $out.tables = @($tables)
    } catch { $errors += "tables: $($_.Exception.Message)" }

    try { $out.shapes = Read-Shapes $d } catch { $errors += "shapes: $($_.Exception.Message)" }

    $paragraphs = @()
    $index = 0

    foreach ($p in $d.Paragraphs) {
        $index += 1
        $record = [ordered]@{ index = ($index - 1) }

        try { $record.text = Get-ParaText $p } catch { $record.text = '' }

        # A paragraph inside a table belongs to the table slice, not this one.
        # It is recorded rather than skipped so an index still means something,
        # but flagged so the node side can leave it out of the body's count.
        try { $record.inTable = [bool]$p.Range.Information(12) } catch { $record.inTable = $null }

        $format = $null
        try { $format = $p.Format } catch { $errors += "format[$index]: $($_.Exception.Message)" }

        if ($null -ne $format) {
            try { $record.alignment = $alignNames[[int]$format.Alignment] } catch { $record.alignment = $null }
            try { $record.leftIndent = Clean-Undefined $format.LeftIndent } catch { $record.leftIndent = $null }
            try { $record.rightIndent = Clean-Undefined $format.RightIndent } catch { $record.rightIndent = $null }
            try { $record.firstLineIndent = Clean-Undefined $format.FirstLineIndent } catch { $record.firstLineIndent = $null }
            try { $record.spaceBefore = Clean-Undefined $format.SpaceBefore } catch { $record.spaceBefore = $null }
            try { $record.spaceAfter = Clean-Undefined $format.SpaceAfter } catch { $record.spaceAfter = $null }
            try { $record.lineSpacing = Clean-Undefined $format.LineSpacing } catch { $record.lineSpacing = $null }
            try { $record.keepWithNext = Clean-Bool $format.KeepWithNext } catch { $record.keepWithNext = $null }
            try { $record.keepTogether = Clean-Bool $format.KeepTogether } catch { $record.keepTogether = $null }
            try { $record.widowControl = Clean-Bool $format.WidowControl } catch { $record.widowControl = $null }
            try { $record.pageBreakBefore = Clean-Bool $format.PageBreakBefore } catch { $record.pageBreakBefore = $null }

            # How the leading was meant, not only what it measures. `exact`
            # clips a line taller than itself and `atLeast` grows it, which is
            # the difference between a document that prints and one that loses
            # a descender -- and the two report the same LineSpacing.
            try {
                $rule = [int]$format.LineSpacingRule
                $ruleNames = @{
                    0 = 'single'; 1 = 'onePointFive'; 2 = 'double';
                    3 = 'atLeast'; 4 = 'exactly'; 5 = 'multiple'
                }
                $record.lineSpacingRule = $ruleNames[$rule]
            } catch { $record.lineSpacingRule = $null }

            # Every stop, with the document's own marked apart from Word's.
            #
            # Word's TabStops collection includes the default stops it puts
            # every half inch across the page, so a paragraph declaring one
            # custom stop reports four. CustomTab is what separates the stop
            # the document asked for from the grid it sits on, and a case that
            # counted the collection would be counting the page, not the file.
            try {
                $stops = @()
                $customCount = 0
                foreach ($stop in $format.TabStops) {
                    $isCustom = $false
                    try { $isCustom = [bool]$stop.CustomTab } catch { }
                    if ($isCustom) { $customCount += 1 }
                    $stops += [ordered]@{
                        position  = [math]::Round([double]$stop.Position, 2)
                        alignment = $tabAlignNames[[int]$stop.Alignment]
                        leader    = $tabLeaderNames[[int]$stop.Leader]
                        custom    = $isCustom
                    }
                }
                $record.tabStops = $stops
                $record.tabStopCount = $stops.Count
                $record.customTabStopCount = $customCount
            } catch {
                $record.tabStops = @()
                $record.tabStopCount = $null
                $record.customTabStopCount = $null
            }

            try {
                $shd = To-Hex $format.Shading.BackgroundPatternColor
                $record.shading = $shd
            } catch { $record.shading = $null }

            # Which edges carry a border. wdLineStyleNone is 0.
            try {
                $record.borders = [ordered]@{
                    top    = ([int]$p.Borders(-1).LineStyle -ne 0)
                    left   = ([int]$p.Borders(-2).LineStyle -ne 0)
                    bottom = ([int]$p.Borders(-3).LineStyle -ne 0)
                    right  = ([int]$p.Borders(-4).LineStyle -ne 0)
                }
            } catch { $record.borders = $null }
        }

        # The run properties of the paragraph's own text.
        #
        # Measured over the text WITHOUT the paragraph mark. The mark carries
        # its own formatting, and Word reports any property that differs across
        # a range as wdUndefined -- so asking the whole paragraph range about a
        # label set in tracked capitals came back "mixed" for both, purely
        # because the invisible character at the end of it is neither.
        try {
            $textRange = $p.Range.Duplicate
            if ($textRange.End -gt $textRange.Start) {
                $null = $textRange.MoveEnd(1, -1)
            }
            $font = $textRange.Font
            $record.fontName = [string]$font.Name
            $record.fontSize = Clean-Undefined $font.Size
            $record.bold = Clean-Bool $font.Bold
            $record.italic = Clean-Bool $font.Italic
            $record.allCaps = Clean-Bool $font.AllCaps
            $record.characterSpacing = Clean-Undefined $font.Spacing
            $record.fontColor = To-Hex $font.Color
        } catch { $errors += "font[$index]: $($_.Exception.Message)" }

        # Where the paragraph actually landed. The start of the range, so the
        # number is where its first character is drawn -- and the end, so a
        # right edge is available to whatever needs one.
        try {
            $start = $p.Range.Duplicate
            $null = $start.Collapse($wdCollapseStart)
            $record.page = [int](Get-Info $start $wdInfoPage)
            $record.x = Get-Info $start $wdInfoX
            $record.y = Get-Info $start $wdInfoY
        } catch { $errors += "start[$index]: $($_.Exception.Message)" }

        try {
            $end = $p.Range.Duplicate
            $null = $end.MoveEnd(1, -1)   # off the paragraph mark
            $null = $end.Collapse($wdCollapseEnd)
            $record.xEnd = Get-Info $end $wdInfoX
            $record.yEnd = Get-Info $end $wdInfoY
            $record.pageEnd = [int](Get-Info $end $wdInfoPage)
        } catch { $errors += "end[$index]: $($_.Exception.Message)" }

        # Whether the paragraph was cut across a page break -- the fact a
        # keep-lines-together case is actually about.
        try {
            if ($null -ne $record.page -and $null -ne $record.pageEnd) {
                $record.linesSplitAcrossPages = ($record.page -ne $record.pageEnd)
            } else {
                $record.linesSplitAcrossPages = $null
            }
        } catch { $record.linesSplitAcrossPages = $null }

        # How many lines Word broke the paragraph into.
        try { $record.lineCount = [int]$p.Range.ComputeStatistics(1) } catch { $record.lineCount = $null }

        $paragraphs += $record
    }

    $out.paragraphs = $paragraphs

    try {
        $d.ExportAsFixedFormat($pdfFull, $wdExportPdf)
        $out.pdfExported = Test-Path -LiteralPath $pdfFull
    } catch { $errors += "pdf: $($_.Exception.Message)" }

    $out.ok = $true
} catch {
    $errors += "fatal: $($_.Exception.Message)"
} finally {
    if ($null -ne $d) {
        try { $d.Close($wdDoNotSaveChanges) } catch { }
    }
    if ($null -ne $w) {
        try { $w.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null } catch { }
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

$out.errors = $errors
$out | ConvertTo-Json -Depth 12 -Compress
