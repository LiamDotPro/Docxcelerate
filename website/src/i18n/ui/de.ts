/**
 * German UI strings.
 *
 * Addresses the reader as "Sie" throughout — the sectors the product is shown
 * working in (Bildung, Wohnen, Versicherung) are institutional, and the copy
 * should not switch register halfway down a page.
 *
 * "Node", "Engine", "Template", "Workspace" and "Helper" stay English: they
 * are the identifiers in the API being described. "Preview" becomes "Vorschau"
 * because it names a thing on screen rather than anything in the code.
 */
import type { UiStrings } from "./en";

export const de: UiStrings = {
  meta: {
    tagline: "Dokumente als Komponenten. DOCX als Ergebnis.",
    description:
      "Ein TypeScript-Toolkit und eine CLI für DOCX-Dokumentprojekte. Schreiben Sie " +
      "Dokumente als typisierte Komponenten, sehen Sie sie im Browser in der Vorschau " +
      "und packen Sie echte .docx-Dateien lokal.",
  },

  common: {
    skipToContent: "Zum Inhalt springen",
    soon: "bald",
  },

  nav: {
    label: "Hauptnavigation",
    docs: "Doku",
    registry: "Registry",
    cloud: "Cloud",
    github: "GitHub",
    switchTheme: "Design wechseln",
    language: "Sprache",
  },

  footer: {
    licence:
      "Dokumente als Komponenten, DOCX als Ausgabe. Offen entwickelt, MIT-lizenziert " +
      "und auf Ihrem eigenen Rechner laufend, solange Sie es nicht anders entscheiden.",
    licenceFile: "MIT-Lizenz lesen",
    docs: "Doku",
    project: "Projekt",
    engine: "Engine",
    cloud: "Cloud",
    github: "GitHub",
    npm: "npm",
    releases: "Releases",
    artifacts: "Build-Artefakte",
    copyright: "© {year} Docxcelerate-Mitwirkende.",
  },

  home: {
    headlineLineOne: "Dokumente als Komponenten.",
    headlineLineTwo: "DOCX als Ergebnis.",
    lead:
      "Setzen Sie Dokumente aus kleinen, typisierten Komponenten zusammen, mit dem " +
      "JSX, das Sie ohnehin schreiben. Binden Sie KI mühelos ein, damit sie Inhalte " +
      "schreibt oder über sie entscheidet. Nutzen Sie unsere Engine, um Dokumente " +
      "in großem Umfang zu erzeugen.",
    ctaStart: "Loslegen →",
    ctaHowItWorks: "So funktioniert es",
    copied: "Kopiert",
    copyFallback: "⌘C drücken",
    integrityTitle: "{hash} · zum Kopieren klicken",
    /** Accessible name for the registry marks beside the version. */
    onNpm: "Docxcelerate auf npm",
    onJsr: "Docxcelerate auf JSR",

    /** The hero's switch between the two builds of the document it shows. */
    builds: {
      label: "Eine Vorlage · zwei Builds",
      engine: "Engine",
      preview: "Vorschau",
      engineCaption: "Engine: Die hervorgehobenen Teile hat ein Modell für diesen Kunden geschrieben, und die Entscheidung hat es selbst getroffen. Jede Zahl wird berechnet, nie generiert.",
      previewCaption: "Vorschau: dieselbe Vorlage auf Ihrem Rechner. Generierte Teile zeigen ihre Platzhalter, nichts wird irgendwohin gesendet, und die Seite sieht bei jedem Speichern gleich aus.",
    },

    /** §1: one paragraph followed from source to preview to the engine. */
    source: {
      title: "Quelle. Vorschau. Dokument.",
      lead: "Ein Absatz der Portfolioübersicht, verfolgt von der Datei, die Sie schreiben, bis zur Seite, die Ihr Kunde öffnet. Die Vorschau wartet nie auf ein Modell, und nichts verlässt Ihren Rechner; die Engine schreibt den echten Satz für jede Person.",
      sourceLabel: "Quelle",
      previewLabel: "Vorschau",
      sourceNote: "Einmal geschrieben, in Ihrem Repo. Der Prompt steht neben dem Node, den er füllt.",
      previewNote: "Bei jedem Speichern neu gebaut. Generierte Nodes zeigen ihren Platzhalter.",
      engineNote: "Für jeden Empfänger geschrieben. Die Sätze drumherum bleiben, wie sie sind.",
    },

    points: {
      authoring: {
        label: "Schreiben",
        title: "Dokumente schreiben wie Websites",
        body:
          "Ein Dokument ist ein Baum aus typisierten Komponenten. Wer schon einmal " +
          "React geschrieben hat, kennt diese Form bereits: Props, Komposition, " +
          "kleine Dateien. Frontend-Entwicklerinnen und -Entwickler sind schon am " +
          "ersten Nachmittag produktiv, statt erst eine Templatesprache lernen " +
          "zu müssen.",
      },
      changeControl: {
        label: "Änderungskontrolle",
        title: "Dokumente liegen in Ihrem Repository",
        body:
          "Weil ein Dokument Quellcode ist, wird das Ändern eines Satzes zu einem " +
          "Pull Request: mit Diff, mit Review, und noch ein Jahr später " +
          "nachvollziehbar, wenn jemand fragt, wer die Formulierung zum " +
          "Zahlungsrückstand geändert hat. Tests halten fest, dass ein Dokument das " +
          "rendert, was Sie erwarten, sodass die CI den Fehler vor der Empfängerin " +
          "oder dem Empfänger bemerkt.",
      },
    },

    pullQuote:
      "Ein Dokument ist ein Baum aus kleinen Komponenten, gerendert von etwas, " +
      "das Papier versteht. Sie bekommen die Ergonomie eines Komponentenmodells. " +
      "Die Person am anderen Ende bekommt eine Word-Datei.",

    attribution: {
      name: "Liam",
      role: "Autor, Docxcelerate",
      portraitAlt: "Liam, der Autor von Docxcelerate",
    },

    engine: {
      eyebrow: "Die Engine",
      title: "Einmal veröffentlichen. Beliebig skalieren.",
      bodyOne:
        "In der Engine werden Dokumente tatsächlich geschrieben. Sie setzt Ihre " +
        "Daten ein, führt die KI aus und gibt das fertige Dokument zurück. Jede Art " +
        "von Node kann KI nutzen, nicht nur Absätze. Die Antwort des Modells wird " +
        "entweder zum Text, geschrieben aus den Informationen, die Sie ihm geben, " +
        "oder trifft eine Entscheidung, von der das Dokument abhängt.",
      bodyTwo:
        "Sie veröffentlichen ein Template einmal in der Engine. Danach kann jedes " +
        "System ihre API mit einem Datensatz aufrufen und bekommt ein Dokument " +
        "zurück. Eine kostenlose Engine steht zum Selbsthosten bereit. Die " +
        "verwaltete Cloud betreibt die vollständige Fassung, mit vielem, das die " +
        "kostenlose nicht hat, und kommt bald.",
      ctaHow: "So arbeitet die Engine →",
      ctaCloud: "Verwaltete Cloud",
      /** {count} is a number of documents. */
      volume: "{count} Übersichten in diesem Quartal",
      steps: {
        build: {
          title: "Bauen",
          detail:
            "Das Framework macht aus Ihrem Dokument ein Paket. Dieser Schritt läuft auf Ihrem Rechner.",
        },
        publish: {
          title: "Veröffentlichen",
          detail:
            "Sie senden das Paket an eine Engine. Die Engine speichert es und gibt ihm einen Namen.",
        },
        write: {
          title: "Schreiben",
          detail:
            "Ihre Anwendung sendet einen Datensatz. Die Engine gibt das fertige Dokument zurück.",
        },
      },
    },



    docs: {
      eyebrow: "Dokumentation",
      title: "Alles ist dokumentiert",
      lead:
        "Jeder Node-Typ, jedes CLI-Flag und jede Datei, die ein Build schreibt. Dazu " +
        "Vorschauen, die der echte Renderer erzeugt hat, sodass nichts auf der Seite " +
        "einen Helper beschreiben kann, den es nicht mehr gibt.",
      ctaAll: "Vollständige Dokumentation →",
      cards: {
        startHere: {
          title: "Hier anfangen",
          blurb: "Einen Workspace anlegen und Ihr erstes Dokument öffnen.",
        },
        documentsAndNodes: {
          title: "Dokumente und Nodes",
          blurb: "Das Komponentenmodell, aus dem ein Dokument gebaut ist.",
        },
        nodeModel: {
          title: "Das Node-Modell",
          blurb: "Jeder Node-Typ, mit einer Vorschau zu jedem.",
        },
        writingNodes: {
          title: "Nodes schreiben",
          blurb: "Die Bausteine, aus denen ein Dokument besteht.",
        },
        cli: {
          title: "CLI-Befehle",
          blurb: "dxcl init, new, node.",
        },
        entrypoints: {
          title: "Package-Entrypoints",
          blurb: "Alles Importierbare und was es exportiert.",
        },
      },
    },

  },

  demo: {
    showFiles: "Workspace-Dateien einblenden",
    hideFiles: "Workspace-Dateien ausblenden",
    workspaceFiles: "Workspace-Dateien",
    resize: "Workspace und Vorschau aufteilen",
    resizeHint: "Ziehen zum Aufteilen · Doppelklick zum Zurücksetzen",
    zoom: "Vorschau-Zoom",
    zoomIn: "Vergrößern",
    zoomOut: "Verkleinern",
    fit: "Einpassen",
    frameTitle: "Das ausgewählte Dokument, gerendert von Docxcelerate",
    caption:
      "Eine echte Dokumentvorschau, live gerendert. Keine Screenshots und keine versteckten Tricks.",
    counts: "{documents} Dokumente · {files} Dateien",
    sectors: {
      education: "Bildung",
      housing: "Wohnen",
      insurance: "Versicherung",
      consulting: "Dienstleistungen",
      finance: "Finanzen",
    },
    documents: {
      "portfolio-review": "Quartalsübersicht zum Portfolio",
      "offer-of-admission": "Zulassungsbescheid",
      "repairs-appointment": "Reparaturtermin",
      "policy-renewal": "Vertragsverlängerung",
      invoice: "Rechnung",
    },
  },

  docs: {
    sidebarLabel: "Dokumentation",
    tocLabel: "Auf dieser Seite",
    onThisPage: "Auf dieser Seite",
    editOnGithub: "Diese Seite auf GitHub bearbeiten ↗",
    viewAsMarkdown: "Diese Seite als Markdown lesen",
    /** {version} is the release the page was built against. */
    renderedAgainst: "Gerendert mit {version}",
    previous: "Zurück",
    next: "Weiter",
    /** Names the previous/next links at the foot of a page. */
    pagination: "Vorherige und nächste Seite",
    untranslatedNote:
      "Diese Seite ist noch nicht übersetzt und wird deshalb auf Englisch angezeigt.",

    groups: {
      "Start Here": "Hier anfangen",
      "Essentials": "Grundlagen",
      "Nodes": "Nodes",
      "CLI": "CLI",
      "Projects": "Projekte",
      "Generation": "Erzeugung",
      "Reference": "Referenz",
    },
    subgroups: {
      "Node types": "Node-Typen",
    },
  },

  nodes: {
    helpers: "Helper",
    noHelpers: "Keine, von Hand geschrieben",
    kind: "Node-Art",
    category: "Kategorie",
    resolves: "Wird aufgelöst",
    children: "Kinder",
    option: "Option",
    type: "Typ",
    whatItDoes: "Was sie bewirkt",
    required: "erforderlich",
    renderNote: "Was die Renderer heute tun",
    resolvesTo: "Wozu es aufgelöst wird",
    resolvesToBody:
      "Der Node, wie er im {model} erscheint: das JSON, das ein Renderer bekommt. " +
      "Kein Styling, kein Layout.",
    endpointAsked: "Was der Endpoint gefragt wird",
    endpointAskedBody:
      "Gegen dieselben Beispieldaten aufgelöst. Ein Vorschau-Build hält beim " +
      "Platzhalter an; ein Build zur Anfragezeit schickt diese mit.",
    previewTitle: "{type}: {variant}, gerendert",
    reference: "{title}-Referenz, mit Vorschauen →",
    categories: {
      Structure: "Struktur",
      Text: "Text",
      Media: "Medien",
      Data: "Daten",
    },
    status: {
      "stable": "Stabil",
      "no-helper": "Noch kein Helper",
      "planned": "Geplant",
    },
  },

  registry: {
    /** The detail pages, per design boards 2b and 2c. */
    backToRegistry: "Registry",
    renderedBy: "Von Docxcelerate beim Build gerendert",
    openPreview: "Öffnen ↗",
    specification: "Spezifikation",
    specNote: "Werte stammen aus style.page und style.typography.",
    margins: "Ränder",
    body: "Fließtext",
    titleLabel: "Titel",
    sectionHeading: "Abschnittsüberschrift",
    paragraphLabel: "Absatz",
    after: "danach",
    compare: "Vergleichen",
    compareNote: "Derselbe Absatz in jedem Theme",
    viewing: "angezeigt",
    requires: "Benötigt",
    requiresNothing: "Nichts. Es installiert allein.",
    installsAlone: "Es installiert allein.",
    stepData: "Vorschaudaten",
    stepDataNote: "das JSON, mit dem diese Vorschau gebaut wurde",
    stepPreview: "Gerenderte Vorschau",
    stepPreviewNote: "die Komponente für sich gerendert",
    stepPreviewArrangementsNote: "jede Anordnung für sich gerendert",
    /** Marks the arrangement a component uses when none is asked for. */
    arrangementDefault: "Standard",
    /** Labels the picker that chooses which arrangement the preview shows. */
    arrangementLabel: "Anordnung",
    /** The picker's option for showing every arrangement at once. */
    arrangementAll: "Alle Anordnungen",
    stepNodes: "Aufgelöste Nodes",
    /** {count} is a number of nodes. */
    stepNodesNote: "was ein Renderer erhält — {count} Nodes",
    stepFiles: "Die Dateien",
    stepFilesNote: "was die Installation schreibt; die gewählte Datei ist die Komponente",
    dataFieldsNote: "Fügen Sie diese nach der Installation Ihrer types.ts hinzu. Das ist der einzige Handgriff.",
    copyFieldPaths: "Feldpfade kopieren",
    dataFields: "Datenfelder",
    copy: "Kopieren",
    copied: "Kopiert",
    copyFile: "{file} kopieren",
    /** {exports} is a list of names, {count} a number of nodes. */
    exportsResolves: "exportiert {exports} · löst zu {count} Nodes auf",
    categoryTags: "Kategorie · Tags",
    nothingHere: "Hier ist noch nichts.",
    nothingHereNote: "Die Registry wird aus dem Paket gebaut. Sobald ein Eintrag ausgeliefert wird, erscheint er hier.",
    /** The one-page registry at /registry. */
    browse: {
      title: "Registry",
      description:
        "Fertige Themes und Dokumentkomponenten für Docxcelerate, mit einem Befehl in ein Projekt installiert.",
      eyebrow: "Registry",
      heading: "Registry",
      lead:
        "Fertige Themes und Komponenten für Ihre Dokumentprojekte. Mit einem einzigen Befehl installieren und dann wie eigenen Code bearbeiten.",
      allTags: "Alle",
      themesNote: "Ein Theme zu installieren schreibt document-style.ts",
      nodes: "Nodes",
      noPreview: "Keine Vorschau",
      /** Leitet die Datenpfade einer Komponentenzeile ein; verbindet Filter-Tags. */
      readsPrefix: "liest",
      /** The page's headline; `heading` is the short name the nav uses. */
      headline: "Fangen Sie mit Papier an, das schon funktioniert.",
      /** {themes} and {components} are counts. */
      counts: "{themes} Themes · {components} Komponenten",
      useTheme: "Theme verwenden",
      selected: "✓ Ausgewählt",
      add: "+ Hinzufügen",
      added: "✓ Hinzugefügt",
      familyLabel: "Was sie zeichnet",
      familyNote: "Gruppiert danach, worin eine Komponente aufgelöst wird, aus dem Build gelesen — kein Etikett, das jemand pflegt.",
      /** {count} is how many themes and components are picked. */
      trayCount: "Ihre Installation · {count} ausgewählt",
      trayEmpty: "ein Theme oder eine Komponente wählen",
    },
    themes: {
      title: "Themes",
      description:
        "Fertige Dokument-Themes für Docxcelerate: Schriften, Farben, Abstände " +
        "und Seitenformat, mit einem Befehl im Projekt.",
      heading: "Themes",
      lead:
        "Schriften, Farben, Abstände und Seite, unter einem Namen. Ein Theme " +
        "ist ein Import im Projekt und reist mit dem gebauten Dokument mit — " +
        "wer es rendert, muss nichts installieren.",
      cta: "Themes ansehen",
    },
    components: {
      title: "Komponenten",
      description:
        "Fertige Docxcelerate-Komponenten: Briefköpfe, Adressblöcke, " +
        "Zahlungsübersichten und mehr, als Quellcode in Ihrem Projekt.",
      heading: "Komponenten",
      lead:
        "Nodes, die jemand schon geschrieben und durchdacht hat. Hinzufügen " +
        "kopiert den Quellcode in Ihr Projekt — keine Version, kein " +
        "Upgrade-Pfad, nichts zu forken.",
      cta: "Komponenten ansehen",
    },
    install: "Installieren",
    preview: "Vorschau",
    category: "Kategorie",
    tags: "Schlagwörter",
    fonts: "Schriften",
    page: "Seite",
    typography: "Typografie",
    colours: "Farben",
    spacing: "Abstände",
    exports: "Exporte",
    field: "Feld",
    type: "Typ",
    whatItDoes: "Wofür es da ist",
    source: "Quellcode",
    lands: "Landet unter",
    resolvesTo: "Was dabei herauskommt",
    previewData: "Vorschau mit",
    drawnAgainst: "Entworfen für das Theme {theme}.",
    allThemes: "Alle Themes",
    allComponents: "Alle Komponenten",
    ad: {
      heading: "Verschaffen Sie Ihrem nächsten Dokument einen Vorsprung.",
      body:
        "Die Registry hält Themes und Dokumentkomponenten bereit. `dxcl add` kopiert die Datei in Ihr Projekt. Keine Abhängigkeit, keine Version zu verfolgen.",
      browse: "Registry durchsehen",
      howItWorks: "Wie das Installieren funktioniert",
      written: "2 Dateien geschrieben · keine Abhängigkeit hinzugefügt",
      afterInstall:
        "Nach der Installation liegen die Dateien in Ihrem Repo. Bearbeiten Sie sie wie jeden anderen Quelltext.",
    },
  },

  notFound: {
    title: "Seite nicht gefunden",
    description: "Diese Seite gibt es nicht.",
    heading: "Diese Seite gibt es nicht",
    body:
      "Der Link ist vielleicht veraltet, oder die Seite ist umgezogen. Die " +
      "Dokumentation ist der beste Ort, um den Faden wieder aufzunehmen.",
    ctaDocs: "Dokumentation lesen →",
    ctaHome: "Zurück zur Startseite",
  },
};
