# Probe C -- Word ground truth for the invoice verification harness.
#
# Drives Word 16 over COM, measures the packed .docx per VERIFY-CONTRACT.md
# ("measure-c.json"), and emits ONE line of compact JSON on stdout. Word is the
# only engine whose page/x/y numbers count as ground truth here, which is why
# this probe exists at all: the file and the preview can both be "right" while
# Word paginates differently.
#
# Every measurement sits in its own try/catch: a region that cannot be located
# records null (never a throw), and a COM hiccup in one step must not cost the
# rest of the board its data. The finally block always closes the document and
# quits Word; the .mjs wrapper kills WINWORD if this script hangs past 180s.
#
# PowerShell 5.1: no && chains, no ternary. The script body is ASCII-only --
# the pound sign and the middle dot are built from [char] codes so the file
# parses identically whether 5.1 reads it as ANSI or UTF-8.

param(
    [Parameter(Mandatory = $true)][string]$DocxPath,
    [Parameter(Mandatory = $true)][string]$PdfPath
)

$ErrorActionPreference = 'Stop'

# PS 5.1 writes redirected stdout in the console codepage, which turns the
# pound sign into an invalid byte for the Node wrapper's UTF-8 decode. Emit
# UTF-8 so the JSON survives the pipe intact.
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

# Word constants, per the contract's environment facts.
$wdInfoPage = 3        # wdActiveEndPageNumber
$wdInfoX = 5           # wdHorizontalPositionRelativeToPage (points)
$wdInfoY = 6           # wdVerticalPositionRelativeToPage (points)
$wdStatPages = 2       # wdStatisticPages
$wdCollapseStart = 1
$wdCollapseEnd = 0
$wdFindStop = 0
$wdCharacter = 1
$wdExportPdf = 17
$wdHeaderPrimary = 1
$wdHeaderFirstPage = 2
$wdDoNotSaveChanges = 0
$wdWithInTable = 12
$wdStatisticLines = 1
$wdBorderBottom = -3
$wdLine = 5
$wdFirstCharacterLineNumber = 10

# Regions whose right edge an objective needs as well as their left: F10 sizes
# the summary label by its width, which is xEnd - x.
$regionsNeedingEnd = @('summary-label')

# Finds an anchor in the given range's story. Case-sensitive first; falls back
# to case-insensitive only when allowed -- an all-caps anchor like PAYMENT must
# NOT fall back, or it would match "Awaiting payment" on page one instead of
# the page-two wordmark.
function Find-InRange {
    param($range, [string]$anchor, [bool]$allowInsensitive)
    $cases = @($true)
    if ($allowInsensitive) { $cases = @($true, $false) }
    foreach ($mc in $cases) {
        try {
            $r = $range.Duplicate
            $f = $r.Find
            $null = $f.ClearFormatting()
            $f.Forward = $true
            $f.Wrap = $wdFindStop
            $f.MatchCase = $mc
            $f.MatchWildcards = $false
            if ($f.Execute($anchor)) { return $r }
        } catch { }
    }
    return $null
}

# Information() with a Select() fallback. Range.Information(5) returns -1 for
# ranges Word has not fully laid out (observed on pages 2-3 of this document);
# selecting the range in the hidden window forces layout, and Selection then
# reports a real position. A value still negative after that is Word saying
# "cannot determine" -- recorded as null, never as a fake coordinate.
function Get-InfoWithFallback {
    param($range, [int]$infoIndex)
    $v = $null
    try { $v = [double]$range.Information($infoIndex) } catch { return $null }
    if ($v -lt 0) {
        try {
            $null = $range.Select()
            $v = [double]$script:w.Selection.Information($infoIndex)
        } catch { }
    }
    if ($v -lt 0) { return $null }
    return [math]::Round($v, 2)
}

# page/x/y of a range's START -- the contract's {page, x, y} shape. Collapsing
# a duplicate keeps the caller's range intact for xEnd / font sampling.
function Get-RangeStartInfo {
    param($range)
    try {
        $s = $range.Duplicate
        $null = $s.Collapse($wdCollapseStart)
        $page = Get-InfoWithFallback $s $wdInfoPage
        if ($null -ne $page) { $page = [int]$page }
        return [ordered]@{
            page = $page
            x    = Get-InfoWithFallback $s $wdInfoX
            y    = Get-InfoWithFallback $s $wdInfoY
        }
    } catch {
        return $null
    }
}

# Cell text without Word's \r\a end-of-cell marker.
function Get-CellText {
    param($cell)
    $t = $cell.Range.Text
    if ($null -eq $t) { return '' }
    return $t.TrimEnd([char]13, [char]7)
}

# One line per \r, control characters stripped -- Word range text is full of
# cell markers (\a) when the story holds a table.
function Get-CleanLines {
    param([string]$text)
    $lines = @()
    if ($null -eq $text) { return , $lines }
    foreach ($piece in $text.Split([char]13)) {
        $clean = ($piece -replace "[`a`v`f`n]", '').Trim()
        if ($clean.Length -gt 0) { $lines += $clean }
    }
    return , $lines
}

$docxFull = (Resolve-Path -LiteralPath $DocxPath).Path
$pdfFull = [System.IO.Path]::GetFullPath($PdfPath)

$out = [ordered]@{
    probe               = 'C'
    ok                  = $false
    docx                = $docxFull
    pdf                 = $pdfFull
    wordVersion         = $null
    pages               = $null
    sections            = $null
    pageSetup           = $null
    regions             = $null
    regionsNote         = 'rule/p2-rule/charges-rule are geometric (no text anchor) and always null in probe C; footer regions fall back to the footer story, marked story=footer'
    amounts             = @()
    firstPageHeaderText = $null
    primaryHeaderText   = $null
    differentFirstPage  = $null
    footer              = $null
    footerTable         = $null
    vAlign              = $null
    suppressedHeadingFinds = $null
    scanCard            = $null
    summary             = $null
    tables              = $null
    chargesTable        = $null
    heading1Style       = $null
    pageNumberParagraph = $null
    fields              = $null
    inlineShapesCount   = $null
    shapesCount         = $null
    terms               = $null
    pdfExported         = $false
    errors              = @()
}
$errors = @()

$w = $null
$d = $null

try {
    $w = New-Object -ComObject Word.Application
    $w.Visible = $false
    $w.DisplayAlerts = 0
    $out.wordVersion = [string]$w.Version

    # Open read-only: FileName, ConfirmConversions:=false, ReadOnly:=true,
    # AddToRecentFiles:=false, ..., Visible:=false. The document must never be
    # touched -- the harness measures, it does not improve.
    $missing = [System.Reflection.Missing]::Value
    $d = $w.Documents.Open($docxFull, $false, $true, $false,
        $missing, $missing, $missing, $missing, $missing, $missing, $missing, $false)

    # Information(5/6) is page-relative only in Print Layout; in another view
    # Word reports column-relative numbers that look plausible and are wrong
    # (observed: two runs of this probe disagreeing by exactly the margins).
    # Force the view so the geometry is deterministic. wdPrintView = 3.
    try {
        $view = $d.ActiveWindow.View
        if ([int]$view.Type -ne 3) { $view.Type = 3 }
    } catch { $errors += "view: $($_.Exception.Message)" }

    try { $out.pages = [int]$d.ComputeStatistics($wdStatPages) } catch { $errors += "pages: $($_.Exception.Message)" }
    try { $out.sections = [int]$d.Sections.Count } catch { $errors += "sections: $($_.Exception.Message)" }

    # Page geometry: F4.C computes footer height from these, and mm objectives
    # need the margins to translate page-relative points into content-relative.
    try {
        $ps = $d.Sections.Item(1).PageSetup
        $out.pageSetup = [ordered]@{
            pageWidth      = [math]::Round([double]$ps.PageWidth, 2)
            pageHeight     = [math]::Round([double]$ps.PageHeight, 2)
            topMargin      = [math]::Round([double]$ps.TopMargin, 2)
            bottomMargin   = [math]::Round([double]$ps.BottomMargin, 2)
            leftMargin     = [math]::Round([double]$ps.LeftMargin, 2)
            rightMargin    = [math]::Round([double]$ps.RightMargin, 2)
            headerDistance = [math]::Round([double]$ps.HeaderDistance, 2)
            footerDistance = [math]::Round([double]$ps.FooterDistance, 2)
        }
        $out.differentFirstPage = [bool]$ps.DifferentFirstPageHeaderFooter
    } catch { $errors += "pageSetup: $($_.Exception.Message)" }

    # --- Regions (contract region table). charges-body-end is the extra key
    # carrying the union's far anchor so F7.C can difference the two y values.
    $mid = [string][char]0xB7
    $anchorTable = [ordered]@{
        'letterhead'       = ('Software consultancy ' + $mid + ' Manchester')
        'band'             = 'Issue date'
        'status-pill'      = 'Awaiting payment'
        'parties'          = 'Billed to'
        'summary-label'    = 'Engagement summary'
        'summary'          = 'Sprint 14 closed out'
        'charges-head'     = 'Description'
        'charges-body'     = 'Discovery and scoping workshop'
        'charges-body-end' = 'Production support retainer'
        'totals-panel'     = 'Subtotal'
        'total-bar'        = 'Total due'
        'closer'           = 'are on page 2'
        'footer'           = 'Registered in England'
        'p2-letterhead'    = 'PAYMENT'
        'bank-grid'        = 'Sort code'
        'reference-panel'  = 'Payment reference'
        'terms'            = 'Payment within 14 days'
        'scan-card'        = 'Scan to pay'
        'footer-2'         = 'Registered in England'
    }

    $regions = [ordered]@{}
    try {
        foreach ($name in $anchorTable.Keys) {
            $anchor = [string]$anchorTable[$name]
            # All-caps anchors stay case-sensitive; see Find-InRange.
            $allowInsensitive = ($anchor -cne $anchor.ToUpperInvariant())
            $hit = Find-InRange $d.Content $anchor $allowInsensitive
            if ($null -ne $hit) {
                $info = Get-RangeStartInfo $hit
                if ($null -ne $info -and $regionsNeedingEnd -contains $name) {
                    $endR = $hit.Duplicate
                    $null = $endR.Collapse($wdCollapseEnd)
                    $info['xEnd'] = Get-InfoWithFallback $endR $wdInfoX
                }
                $regions[$name] = $info
            } else {
                $regions[$name] = $null
            }
        }
    } catch { $errors += "regions: $($_.Exception.Message)" }

    # Footer text lives in the footer story, which Content.Find cannot reach;
    # fall back there so the footer regions still get geometry.
    foreach ($name in @('footer', 'footer-2')) {
        if ($null -eq $regions[$name]) {
            try {
                $fr = $d.Sections.Item(1).Footers.Item($wdHeaderPrimary).Range
                $hit = Find-InRange $fr ([string]$anchorTable[$name]) $true
                if ($null -ne $hit) {
                    $info = Get-RangeStartInfo $hit
                    if ($null -ne $info) { $info['story'] = 'footer' }
                    $regions[$name] = $info
                }
            } catch { $errors += "${name}: $($_.Exception.Message)" }
        }
    }

    # Geometric strips have no text to find; explicit nulls keep the key set
    # stable for objectives.mjs.
    foreach ($name in @('rule', 'p2-rule', 'charges-rule')) { $regions[$name] = $null }
    $out.regions = $regions

    # --- Amount column: every currency string in the built model, located with
    # xEnd (range collapsed to END) because F6.C asserts on the RIGHT edge --
    # tabular figures line up on the right or not at all.
    try {
        $amounts = @()
        $modelPath = Join-Path (Split-Path -Parent $docxFull) 'model.json'
        if (Test-Path -LiteralPath $modelPath) {
            $pound = [string][char]0xA3
            $rx = [regex]($pound + '[\d,]+\.\d\d')
            $seen = @{}
            foreach ($m in $rx.Matches((Get-Content -Raw -Encoding UTF8 -LiteralPath $modelPath))) {
                if (-not $seen.ContainsKey($m.Value)) {
                    $seen[$m.Value] = $true
                    $entry = [ordered]@{ text = $m.Value; page = $null; x = $null; y = $null; xEnd = $null; tableCell11 = $null; columnIndex = $null; columnCount = $null }
                    $hit = Find-InRange $d.Content $m.Value $false
                    if ($null -ne $hit) {
                        $info = Get-RangeStartInfo $hit
                        if ($null -ne $info) {
                            $entry.page = $info.page
                            $entry.x = $info.x
                            $entry.y = $info.y
                        }
                        $endR = $hit.Duplicate
                        $null = $endR.Collapse($wdCollapseEnd)
                        $entry.xEnd = Get-InfoWithFallback $endR $wdInfoX
                        # Which table this figure sits in, named by its header
                        # cell. The charge amounts and the totals panel are
                        # different columns, so an objective about "the amount
                        # column" has to be able to tell them apart.
                        try {
                            if ($hit.Information($wdWithInTable)) {
                                $ownerTable = $hit.Tables.Item(1)
                                $ownerText = Get-CellText $ownerTable.Cell(1, 1)
                                if ($ownerText.Length -gt 40) { $ownerText = $ownerText.Substring(0, 40) }
                                $entry.tableCell11 = $ownerText
                                # Which column, so an objective about "the
                                # amount column" is not handed the rate column
                                # as well -- three columns of figures line up
                                # on three different edges, by design.
                                $entry.columnIndex = [int]$hit.Cells.Item(1).ColumnIndex
                                $entry.columnCount = [int]$ownerTable.Columns.Count
                            }
                        } catch { }
                    }
                    $amounts += $entry
                }
            }
        } else {
            $errors += "amounts: model.json not found beside the docx ($modelPath)"
        }
        $out.amounts = @($amounts)
    } catch { $errors += "amounts: $($_.Exception.Message)" }

    # --- Header and footer text (F2.C asserts on these).
    try {
        $out.firstPageHeaderText = ((Get-CleanLines $d.Sections.Item(1).Headers.Item($wdHeaderFirstPage).Range.Text) -join ' | ')
    } catch { $errors += "firstPageHeaderText: $($_.Exception.Message)" }
    try {
        $out.primaryHeaderText = ((Get-CleanLines $d.Sections.Item(1).Headers.Item($wdHeaderPrimary).Range.Text) -join ' | ')
    } catch { $errors += "primaryHeaderText: $($_.Exception.Message)" }

    try {
        $fr = $d.Sections.Item(1).Footers.Item($wdHeaderPrimary).Range
        $footerLines = @()
        # Count paragraphs with visible text rather than raw paragraph marks:
        # a table-borne footer is full of empty structural paragraphs.
        foreach ($p in $fr.Paragraphs) {
            $clean = (($p.Range.Text) -replace "[`r`a`v]", '').Trim()
            if ($clean.Length -gt 0) { $footerLines += $clean }
        }
        # How many lines the strip occupies, which is not how many paragraphs
        # it holds: three cells side by side in one row are three paragraphs
        # and one line. Word reports no line number inside a footer story, so
        # the count comes from the shape -- a table's rows stack, the cells
        # within a row do not, and a paragraph outside any table is its own
        # line.
        $visualLines = 0
        try {
            foreach ($p in $fr.Paragraphs) {
                $clean = (($p.Range.Text) -replace "[`r`a`v]", '').Trim()
                if ($clean.Length -eq 0) { continue }
                if (-not $p.Range.Information($wdWithInTable)) { $visualLines += 1 }
            }
            foreach ($tbl in $fr.Tables) { $visualLines += [int]$tbl.Rows.Count }
        } catch { $visualLines = $null }
        # Height: the footer runs from the top of its first line down to where
        # the footer band ends, which is FooterDistance up from the page edge.
        # Measuring the top by geometry rather than by summing line heights
        # keeps a table-borne footer honest -- its structural paragraphs have
        # heights that the printed strip does not.
        $heightPt = $null
        try {
            $topR = $fr.Duplicate
            $null = $topR.Collapse($wdCollapseStart)
            $topY = Get-InfoWithFallback $topR $wdInfoY
            if ($null -ne $topY -and $null -ne $out.pageSetup) {
                $bandBottom = [double]$out.pageSetup.pageHeight - [double]$out.pageSetup.footerDistance
                $heightPt = [math]::Round($bandBottom - [double]$topY, 2)
            }
        } catch { }
        $out.footer = [ordered]@{
            text           = ((Get-CleanLines $fr.Text) -join ' | ')
            lines          = @($footerLines)
            paragraphCount = $footerLines.Count
            lineCount      = $visualLines
            heightPt       = $heightPt
        }
    } catch { $errors += "footer: $($_.Exception.Message)" }

    # --- The footer's own table (F3.C: edge to edge, 595.3pt wide at x=0).
    try {
        $ffr = $d.Sections.Item(1).Footers.Item($wdHeaderPrimary).Range
        if ($ffr.Tables.Count -gt 0) {
            $ft = $ffr.Tables.Item(1)
            $ftStart = $ft.Range.Duplicate
            $null = $ftStart.Collapse($wdCollapseStart)
            # LeftIndent is the table's own offset from the text column, which
            # is the fact F3 is about. Information(5) reports where the first
            # cell's *text* starts, and that carries the cell's inset -- a
            # table flush to the paper still reads a few points in.
            $leftIndent = $null
            try { $leftIndent = [math]::Round([double]$ft.Rows.LeftIndent, 2) } catch { }
            # How wide the table actually is, added up across its first row.
            #
            # PreferredWidth is the table's *setting*, and Word reports it as
            # wdUndefined for a multi-column table laid out fixed -- the widths
            # are the columns' then, not a preference the table expresses. The
            # cells are where the number lives, and adding them is the same
            # reading whichever layout the table is in.
            $width = $null
            try {
                $total = 0
                foreach ($cell in $ft.Rows.Item(1).Cells) { $total += [double]$cell.Width }
                $width = [math]::Round($total, 2)
            } catch { }
            $out.footerTable = [ordered]@{
                width              = $width
                preferredWidth     = [math]::Round([double]$ft.PreferredWidth, 2)
                preferredWidthType = [int]$ft.PreferredWidthType
                rows               = [int]$ft.Rows.Count
                columns            = [int]$ft.Columns.Count
                leftIndent         = $leftIndent
                x                  = Get-InfoWithFallback $ftStart $wdInfoX
                y                  = Get-InfoWithFallback $ftStart $wdInfoY
            }
        } else {
            $out.footerTable = $null
        }
    } catch { $errors += "footerTable: $($_.Exception.Message)" }

    # --- The three headings F1 suppresses. A Find that misses is the whole
    # objective, so the miss is recorded as a fact rather than an absence.
    try {
        $suppressed = [ordered]@{}
        foreach ($heading in @('Invoice details', 'Parties', 'Charges')) {
            $hit = Find-InRange $d.Content $heading $false
            $suppressed[$heading] = ($null -ne $hit)
        }
        $out.suppressedHeadingFinds = $suppressed
    } catch { $errors += "suppressedHeadingFinds: $($_.Exception.Message)" }

    # --- The scan-to-pay card: the table holding the QR image (F13.C, D14.C).
    try {
        $out.scanCard = $null
        # The card is whatever table holds the picture, so it is looked for by
        # the picture -- either the real one or, while the image is still
        # unresolved, the placeholder line standing in for it. "Scan to pay" on
        # its own now finds the section heading, which sits outside the card.
        # The card is the picture in a table of its own -- one row, one cell.
        # The letterhead's mark is also a picture in a table, three columns
        # wide with the sender's name beside it, and is not a card.
        $hit = $null
        foreach ($shape in $d.InlineShapes) {
            try {
                if (-not $shape.Range.Information($wdWithInTable)) { continue }
                $owner = $shape.Range.Tables.Item(1)
                if ([int]$owner.Rows.Count -eq 1 -and [int]$owner.Columns.Count -eq 1) {
                    $hit = $shape.Range
                    break
                }
            } catch { }
        }
        if ($null -eq $hit) { $hit = Find-InRange $d.Content '[image: Scan to pay' $true }
        if ($null -ne $hit -and $hit.Information($wdWithInTable)) {
            $card = $hit.Tables.Item(1)
            $cardStart = $card.Range.Duplicate
            $null = $cardStart.Collapse($wdCollapseStart)
            $cardEnd = $card.Range.Duplicate
            $null = $cardEnd.Collapse($wdCollapseEnd)
            $yTop = Get-InfoWithFallback $cardStart $wdInfoY
            $yBottom = Get-InfoWithFallback $cardEnd $wdInfoY
            $cardHeight = $null
            if ($null -ne $yTop -and $null -ne $yBottom) {
                $cardHeight = [math]::Round([double]$yBottom - [double]$yTop, 2)
            }
            $out.scanCard = [ordered]@{
                preferredWidth    = [math]::Round([double]$card.PreferredWidth, 2)
                rows              = [int]$card.Rows.Count
                columns           = [int]$card.Columns.Count
                x                 = Get-InfoWithFallback $cardStart $wdInfoX
                y                 = $yTop
                heightPt          = $cardHeight
                inlineShapeCount  = [int]$card.Range.InlineShapes.Count
            }
        }
    } catch { $errors += "scanCard: $($_.Exception.Message)" }

    # --- The engagement summary's measure (F14.C: a 20mm right indent, and no
    # line running past 158mm). Word will not hand over a line's width, so each
    # line is walked with the selection: home, read x, end, read x.
    try {
        $out.summary = $null
        $hit = Find-InRange $d.Content 'Sprint 14 closed out' $true
        if ($null -ne $hit) {
            $para = $hit.Paragraphs.Item(1)
            # The measure is how far right the text actually reaches, so it is
            # taken from the words themselves rather than by walking lines with
            # the selection -- Word will not report a position for a range it
            # has not laid out, and the selection-walk loses its place the
            # moment one of those comes back empty.
            $head = $para.Range.Duplicate
            $null = $head.Collapse($wdCollapseStart)
            $leftX = Get-InfoWithFallback $head $wdInfoX
            $maxRight = $null
            $lineCount = $null
            try {
                foreach ($word in $para.Range.Words) {
                    try {
                        $wordEnd = $word.Duplicate
                        $null = $wordEnd.Collapse($wdCollapseEnd)
                        $x = Get-InfoWithFallback $wordEnd $wdInfoX
                        if ($null -ne $x) {
                            if ($null -eq $maxRight -or [double]$x -gt [double]$maxRight) { $maxRight = $x }
                        }
                    } catch { }
                }
                $tail = $para.Range.Duplicate
                $null = $tail.MoveEnd($wdCharacter, -1)
                $null = $tail.Collapse($wdCollapseEnd)
                $firstLine = Get-InfoWithFallback $head $wdFirstCharacterLineNumber
                $lastLine = Get-InfoWithFallback $tail $wdFirstCharacterLineNumber
                if ($null -ne $firstLine -and $null -ne $lastLine) {
                    $lineCount = [int]([double]$lastLine - [double]$firstLine) + 1
                }
            } catch { }
            $maxLine = $null
            if ($null -ne $maxRight -and $null -ne $leftX) {
                $maxLine = [math]::Round([double]$maxRight - [double]$leftX, 2)
            }
            $out.summary = [ordered]@{
                rightIndent  = [math]::Round([double]$para.Format.RightIndent, 2)
                leftIndent   = [math]::Round([double]$para.Format.LeftIndent, 2)
                leftX        = $leftX
                rightmostX   = $maxRight
                lineCount    = $lineCount
                maxLineWidth = $maxLine
            }
        }
    } catch { $errors += "summary: $($_.Exception.Message)" }

    # --- Tables. Word counts top-level tables only, which is what F-objectives
    # index by; the charges table is identified by content, not position, so a
    # reordered document cannot silently shift which table gets asserted on.
    try {
        $tableCount = [int]$d.Tables.Count
        $tableList = @()
        $charges = $null
        $chargesIndex = $null
        $band = $null
        for ($i = 1; $i -le $tableCount; $i++) {
            $t = $d.Tables.Item($i)
            $entry = [ordered]@{ index = $i; preferredWidth = $null; preferredWidthType = $null; rows = $null; columns = $null; cell11 = $null }
            try { $entry.preferredWidth = [math]::Round([double]$t.PreferredWidth, 2) } catch { }
            try { $entry.preferredWidthType = [int]$t.PreferredWidthType } catch { }
            try { $entry.rows = [int]$t.Rows.Count } catch { }
            try { $entry.columns = [int]$t.Columns.Count } catch { }
            try {
                $c11 = Get-CellText $t.Cell(1, 1)
                if ($c11.Length -gt 40) { $c11 = $c11.Substring(0, 40) }
                $entry.cell11 = $c11
                if ($null -eq $charges) {
                    # Word reports display-cased text for a w:caps run, so the
                    # header cell reads DESCRIPTION -- match case-insensitively.
                    if ($c11.ToUpperInvariant().StartsWith('DESCRIPTION')) {
                        $charges = $t
                        $chargesIndex = $i
                    }
                }
                # The band of dates, named the same way and for the same
                # reason: F11 asks about its cells, and an index would follow
                # the wrong table the moment the document is reordered.
                if ($null -eq $band) {
                    if ($c11.ToUpperInvariant().StartsWith('ISSUE DATE')) { $band = $t }
                }
            } catch { }
            $tableList += $entry
        }
        $out.tables = [ordered]@{ count = $tableCount; list = @($tableList) }
    } catch { $errors += "tables: $($_.Exception.Message)" }

    # --- Charges table detail: shading per row (F8.C), vertical alignment
    # (F11.C), and the amount cell's face (F6.C).
    if ($null -ne $charges) {
        try {
            $shading = @()
            for ($r = 1; $r -le 8; $r++) {
                try { $shading += [int64]$charges.Cell($r, 1).Shading.BackgroundPatternColor } catch { $shading += $null }
            }
            $cell11VAlign = $null
            try { $cell11VAlign = [int]$charges.Cell(1, 1).VerticalAlignment } catch { }
            # Body row 2, last cell = the first amount cell. MoveEnd -1 drops
            # the end-of-cell marker, whose font is not the run's.
            $amountFontName = $null
            $amountSample = $null
            try {
                $row = $charges.Rows.Item(2)
                $lastCell = $row.Cells.Item($row.Cells.Count)
                $rr = $lastCell.Range.Duplicate
                $null = $rr.MoveEnd($wdCharacter, -1)
                $amountFontName = [string]$rr.Font.Name
                $amountSample = Get-CellText $lastCell
            } catch { }
            # Lines per charge row (F7.C2: nothing wraps to a third line at
            # the fixture's data lengths). Column 1 carries the description and
            # its muted note, so it is the column that decides a row's height.
            #
            # ComputeStatistics(wdStatisticLines) answers 0 for a cell range,
            # so the count comes from the line numbers Word reports at the
            # cell's first and last character instead.
            $rowLines = @()
            try {
                for ($r = 2; $r -le [int]$charges.Rows.Count; $r++) {
                    $count = $null
                    try {
                        $cellRange = $charges.Cell($r, 1).Range.Duplicate
                        $null = $cellRange.MoveEnd($wdCharacter, -1)
                        $head = $cellRange.Duplicate
                        $null = $head.Collapse($wdCollapseStart)
                        $tail = $cellRange.Duplicate
                        $null = $tail.Collapse($wdCollapseEnd)
                        $first = Get-InfoWithFallback $head $wdFirstCharacterLineNumber
                        $last = Get-InfoWithFallback $tail $wdFirstCharacterLineNumber
                        if ($null -ne $first -and $null -ne $last) {
                            $count = [int]([double]$last - [double]$first) + 1
                        }
                    } catch { }
                    $rowLines += $count
                }
            } catch { }
            # F9.C: the rule under row 3 should not be drawn at all.
            $row3Bottom = $null
            try { $row3Bottom = [int]$charges.Cell(3, 1).Borders.Item($wdBorderBottom).LineStyle } catch { }
            $out.chargesTable = [ordered]@{
                index                 = $chargesIndex
                shadingCol1Rows1to8   = @($shading)
                cell11VerticalAlignment = $cell11VAlign
                amountFontName        = $amountFontName
                amountSampleText      = $amountSample
                rowLineCounts         = @($rowLines)
                row3BottomLineStyle   = $row3Bottom
            }
        } catch { $errors += "chargesTable: $($_.Exception.Message)" }
    } else {
        $errors += 'chargesTable: no table with Cell(1,1) starting "Description"'
    }

    # --- Vertical alignment of the band and footer cells (F11.C). Both tables
    # are named by their header cell rather than by index, so reordering the
    # document cannot quietly point this at a different table.
    try {
        $bandVAlign = @()
        $footerVAlign = @()
        if ($null -ne $band) {
            foreach ($cell in $band.Range.Cells) {
                try { $bandVAlign += [int]$cell.VerticalAlignment } catch { $bandVAlign += $null }
            }
        }
        $ffr = $d.Sections.Item(1).Footers.Item($wdHeaderPrimary).Range
        if ($ffr.Tables.Count -gt 0) {
            foreach ($cell in $ffr.Tables.Item(1).Range.Cells) {
                try { $footerVAlign += [int]$cell.VerticalAlignment } catch { $footerVAlign += $null }
            }
        }
        $out.vAlign = [ordered]@{ band = @($bandVAlign); footer = @($footerVAlign) }
    } catch { $errors += "vAlign: $($_.Exception.Message)" }

    # --- Heading 1 style (F10.C: tracking arrives as Font.Spacing in points).
    try {
        $h1 = $d.Styles.Item('Heading 1')
        $out.heading1Style = [ordered]@{
            fontSpacing = [math]::Round([double]$h1.Font.Spacing, 2)
            fontSize    = [math]::Round([double]$h1.Font.Size, 2)
        }
    } catch { $errors += "heading1Style: $($_.Exception.Message)" }

    # --- The page-number paragraph (F5.C: Alignment 2 = right, SpaceAfter 0).
    # The PAGE field lives in the footer's last paragraph today.
    try {
        $fr = $d.Sections.Item(1).Footers.Item($wdHeaderPrimary).Range
        # The paragraph that actually holds the PAGE field, not the story's
        # last one -- a table-borne footer ends with the empty paragraph Word
        # keeps after every table, which is set like nothing in particular.
        $fieldPara = $null
        foreach ($fld in $fr.Fields) {
            try {
                if ((($fld.Code.Text) -replace '\s+', ' ').Trim() -match '\bPAGE\b') {
                    $fieldPara = $fld.Result.Paragraphs.Item(1)
                    break
                }
            } catch { }
        }
        if ($null -eq $fieldPara) { $fieldPara = $fr.Paragraphs.Item($fr.Paragraphs.Count) }
        $out.pageNumberParagraph = [ordered]@{
            alignment  = [int]$fieldPara.Format.Alignment
            spaceAfter = [math]::Round([double]$fieldPara.Format.SpaceAfter, 2)
        }
    } catch { $errors += "pageNumberParagraph: $($_.Exception.Message)" }

    # --- Fields. Update everything, then read the footer story back. Word
    # evaluates a story once, not per rendered page, so PAGE reads as its
    # first-page value here; the per-page rendering is checked visually via
    # the exported PDF (probe V), and the field codes are recorded so an
    # objective can at least assert PAGE and NUMPAGES both exist.
    try {
        $null = $d.Fields.Update()
        $fr = $d.Sections.Item(1).Footers.Item($wdHeaderPrimary).Range
        $null = $fr.Fields.Update()
        $fieldList = @()
        foreach ($fld in $fr.Fields) {
            $code = $null
            $result = $null
            try { $code = (($fld.Code.Text) -replace '\s+', ' ').Trim() } catch { }
            try { $result = (($fld.Result.Text) -replace "[`r`a`v]", '').Trim() } catch { }
            $fieldList += [ordered]@{ code = $code; result = $result }
        }
        $out.fields = [ordered]@{
            updated               = $true
            footerTextAfterUpdate = ((Get-CleanLines $fr.Text) -join ' | ')
            footerFields          = @($fieldList)
            note                  = 'field results are story-level; per-page footer text is only visible in the exported PDF'
        }
    } catch { $errors += "fields: $($_.Exception.Message)" }

    try { $out.inlineShapesCount = [int]$d.InlineShapes.Count } catch { $errors += "inlineShapesCount: $($_.Exception.Message)" }
    try { $out.shapesCount = [int]$d.Shapes.Count } catch { $errors += "shapesCount: $($_.Exception.Message)" }

    # --- Terms paragraphs: face size and ink of the first 10 characters
    # (enough to sample the run without crossing into a differently-set run).
    try {
        $termsOut = [ordered]@{}
        $termAnchors = [ordered]@{
            paymentWithin14Days  = 'Payment within 14 days'
            sendRemittanceAdvice = 'Send remittance advice'
        }
        foreach ($key in $termAnchors.Keys) {
            $hit = Find-InRange $d.Content ([string]$termAnchors[$key]) $true
            if ($null -ne $hit) {
                $sample = $hit.Duplicate
                $null = $sample.Collapse($wdCollapseStart)
                $sample.End = $sample.Start + 10
                $info = Get-RangeStartInfo $hit
                $termsOut[$key] = [ordered]@{
                    page      = $info.page
                    fontSize  = [math]::Round([double]$sample.Font.Size, 2)
                    fontColor = [int64]$sample.Font.Color
                }
            } else {
                $termsOut[$key] = $null
            }
        }
        $out.terms = $termsOut
    } catch { $errors += "terms: $($_.Exception.Message)" }

    # --- PDF export feeds probe V. A stale PDF must not fake a fresh export.
    try {
        $pdfDir = Split-Path -Parent $pdfFull
        if (-not (Test-Path -LiteralPath $pdfDir)) { $null = New-Item -ItemType Directory -Force -Path $pdfDir }
        if (Test-Path -LiteralPath $pdfFull) { Remove-Item -LiteralPath $pdfFull -Force -Confirm:$false }
        $d.ExportAsFixedFormat($pdfFull, $wdExportPdf)
        $out.pdfExported = [bool](Test-Path -LiteralPath $pdfFull)
    } catch { $errors += "pdfExport: $($_.Exception.Message)" }

    $out.ok = $true
} catch {
    $errors += "fatal: $($_.Exception.Message)"
} finally {
    # Always leave the machine clean: no dirty-save (Close 0), no Word left
    # behind. Failures here are swallowed -- the wrapper's WINWORD kill is the
    # backstop of last resort.
    try { if ($null -ne $d) { $d.Close($wdDoNotSaveChanges) } } catch { }
    try { if ($null -ne $w) { $w.Quit() } } catch { }
    try { if ($null -ne $d) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($d) } } catch { }
    try { if ($null -ne $w) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) } } catch { }
    $d = $null
    $w = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$out.errors = @($errors)
Write-Output (ConvertTo-Json -InputObject $out -Compress -Depth 8)
if (-not $out.ok) { exit 1 }
exit 0
