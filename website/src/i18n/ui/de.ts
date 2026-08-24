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
    themes: "Themes",
    components: "Komponenten",
    cloud: "Cloud",
    github: "GitHub",
    switchTheme: "Design wechseln",
    language: "Sprache",
  },

  footer: {
    licence:
      "Dokumente als Komponenten, DOCX als Ausgabe. Offen entwickelt, MIT-lizenziert " +
      "und auf Ihrem eigenen Rechner laufend, solange Sie es nicht anders entscheiden.",
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
    copyInstall: "Installationsbefehl kopieren",
    copy: "Kopieren",
    copied: "Kopiert",
    copyFallback: "⌘C drücken",
    integrityTitle: "{hash} · zum Kopieren klicken",
    /** Accessible name for the registry marks beside the version. */
    onNpm: "Docxcelerate auf npm",
    onJsr: "Docxcelerate auf JSR",

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
      ai: {
        label: "KI",
        title: "KI auf Komponentenebene",
        body:
          "KI kommt über Hooks herein, in den Komponenten, die Sie ohnehin " +
          "schreiben. Eine Komponente gibt dem Modell ihren Kontext und das, was " +
          "geschrieben werden soll, damit es genau diesen einen Teil des Dokuments " +
          "erzeugt, während alles darum herum deterministisch bleibt. Wie viel " +
          "erzeugt wird, entscheiden Sie, Komponente für Komponente.",
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

    scale: {
      eyebrow: "Enterprise",
      title: "Gebaut für die Berichte, die Sie heute von Hand schreiben",
      lead:
        "Ein Unternehmen, das denselben Bericht hunderte Male im Monat erstellt, " +
        "hat das Layout und die meisten Formulierungen längst. Was sich ändert, ist " +
        "die Person, für die er geschrieben wird. Bauen Sie diesen Bericht aus " +
        "Komponenten neu, halten Sie jeden Durchlauf identisch, und geben Sie einem " +
        "Modell nur die Teile, die davon abhängen, wer ihn liest.",
      traits: {
        volume: {
          title: "Ein Template, jede Empfängerin und jeder Empfänger",
          body:
            "Das Template einmal veröffentlichen, dann pro Person aufrufen. Ein " +
            "Bericht oder hunderttausend sind derselbe Aufruf, wiederholt.",
        },
        determinism: {
          title: "Jedes Mal dasselbe Dokument",
          body:
            "Alles, was Sie nicht als generiert markiert haben, rendert bei jedem " +
            "Lauf identisch. Variieren dürfen nur die Teile, die Sie ausgewählt haben.",
        },
        integration: {
          title: "Aufgerufen von den Systemen, die Sie betreiben",
          body:
            "Ihr CRM, Ihre Fallverwaltung oder Ihr Abrechnungssystem schickt seine " +
            "Daten und bekommt eine .docx zurück. Niemand exportiert eine Tabelle " +
            "oder öffnet Word.",
        },
        reproducibility: {
          title: "Noch ein Jahr später reproduzierbar",
          body:
            "Templates sind versioniert, jedes Dokument lässt sich also aus " +
            "demselben Template und denselben Daten neu bauen. Eine Prüfung bekommt " +
            "einen Build, kein Archiv.",
        },
      },
    },

    openSource: {
      eyebrow: "Open Source",
      title: "Open Source, und darauf angelegt, es zu bleiben",
      bodyOne:
        "Das Framework, die Renderer, das Node-Modell und die CLI sind MIT-lizenziert " +
        "und werden offen entwickelt. Lesen Sie den Code, der Ihre Dokumente " +
        "schreibt, forken Sie ihn oder nehmen Sie ihn in Ihren eigenen Build auf.",
      bodyTwo:
        "Die Engine lässt sich kostenlos selbst hosten. Dokumente im großen Maßstab " +
        "zu erzeugen hängt damit nie daran, dass ein Anbieter am Markt bleibt oder " +
        "eine Preisliste gleich bleibt. Unsere kostenpflichtige Cloud legt die " +
        "Premium-Funktionen auf denselben freien Kern, sodass Hosting und Skalierung " +
        "ab dem ersten Dokument bereitstehen. Sie ist die Bequemlichkeit, nicht der " +
        "Weg hinein.",
      ctaSource: "Zum Quellcode →",
      ctaNpm: "Auf npm ansehen",

      licence: {
        pages: "1 von 1",
        copyright: "Copyright (c) {year} Docxcelerate",
        fork: "Forken",
        vendor: "Einbetten",
        ship: "Ausliefern",
        read: "Lesen →",
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

    agentSkill: {
      eyebrow: "Agent Skills",
      title: "Geben Sie Ihrem Agenten das Ganze",
      lead:
        "Eine einzige Markdown-Datei bringt einem Coding-Agenten bei, wie Dokumente " +
        "hier aufgebaut sind: das Komponentenmodell, die Regeln, über die Agenten " +
        "stolpern, und jeder Befehl. Legen Sie sie ab und fragen Sie nach einem " +
        "Dokument, statt das erste selbst zu schreiben.",
      where: "Wohin die Datei gehört",
      copySkill: "Skill kopieren",
      references:
        "Es ist schlichtes Markdown, und vier Referenzdateien liegen daneben in " +
        "skills/docxcelerate/.",
      agents: {
        "claude-code":
          "Kopieren Sie den Ordner nach .claude/skills/ für ein Projekt oder nach " +
          "~/.claude/skills/ für alle. Er lädt sich selbst, sobald ein " +
          "Dokumentprojekt auftaucht.",
        "cursor":
          "Als Rules-Datei ablegen, dann liest Cursor sie in diesem Projekt. Oder die " +
          "Datei im Chat per @ erwähnen, wenn Sie sie brauchen.",
        "copilot":
          "In .github/copilot-instructions.md einfügen, dann wendet Copilot sie im " +
          "ganzen Repository an.",
        "agents-md":
          "Codex, Gemini CLI, Aider und Cline lesen alle AGENTS.md im Wurzelverzeichnis " +
          "des Repos. Fügen Sie sie ein oder verweisen Sie darauf, wenn Ihr Agent " +
          "Dateien öffnet.",
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
    },
    documents: {
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
      filterByTag: "Nach Tag filtern",
      allTags: "Alle",
      themesNote: "Ein Theme zu installieren schreibt document-style.ts",
      componentsNote: "Eine Komponente zu installieren schreibt ihre Datei nach nodes/",
      nodes: "Nodes",
      noPreview: "Keine Vorschau",
      /** Leitet die Datenpfade einer Komponentenzeile ein; verbindet Filter-Tags. */
      readsPrefix: "liest",
      andJoin: "und",
      /** {tag} is the tag that matched nothing. */
      noMatches: "Keine Einträge passen zu {tag}.",
      noMatchesTitle: "Keine Treffer",
      noMatchesNote: "Probieren Sie ein anderes Tag oder setzen Sie den Filter zurück.",
      clearFilter: "Filter zurücksetzen",
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
