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
    cloud: "Cloud",
    github: "GitHub",
    switchTheme: "Design wechseln",
    language: "Sprache",
  },

  footer: {
    label: "Fußzeile",
    licence:
      "MIT-lizenziert. Schreiben, Vorschau und das Packen von DOCX laufen vollständig auf Ihrem Rechner.",
    docs: "Doku",
    cloud: "Cloud",
    github: "GitHub",
    npm: "npm",
    copyright: "© {year} Docxcelerate-Mitwirkende.",
  },

  home: {
    headlineLineOne: "Dokumente als Komponenten.",
    headlineLineTwo: "DOCX als Ergebnis.",
    lead:
      "Setzen Sie Dokumente aus kleinen, typisierten Komponenten zusammen — mit dem " +
      "JSX, das Sie ohnehin schreiben. Nutzen Sie die eingebauten KI-Funktionen für " +
      "dynamische Dokumente und skalieren Sie die Dokumenterzeugung mit unserer Engine.",
    ctaStart: "Loslegen →",
    ctaHowItWorks: "So funktioniert es",
    copyInstall: "Installationsbefehl kopieren",
    copy: "Kopieren",
    copied: "Kopiert",
    copyFallback: "⌘C drücken",
    integrityTitle: "{hash} — zum Kopieren klicken",

    points: {
      authoring: {
        label: "Schreiben",
        title: "Dokumente schreiben wie Websites",
        body:
          "Ein Dokument ist ein Baum aus typisierten Komponenten. Wer schon einmal " +
          "React geschrieben hat, kennt diese Form bereits — Props, Komposition, " +
          "kleine Dateien — und ist als Frontend-Entwicklerin oder -Entwickler " +
          "schon am ersten Nachmittag produktiv, statt erst eine Templatesprache " +
          "lernen zu müssen.",
      },
      ai: {
        label: "KI",
        title: "Für KI entworfen",
        body:
          "Generierter Text ist ein Node-Typ, kein nachträglicher Aufsatz. " +
          "Markieren Sie den einen Absatz, der wirklich ein Modell braucht, geben " +
          "Sie ihm Prompts und einen Platzhalter, und lassen Sie den Rest " +
          "deterministisch — so löst KI den schwierigen Teil, ohne dass das ganze " +
          "Dokument ihr ausgeliefert ist.",
      },
      changeControl: {
        label: "Änderungskontrolle",
        title: "Dokumente liegen in Ihrem Repository",
        body:
          "Weil ein Dokument Quellcode ist, wird das Ändern eines Satzes zu einem " +
          "Pull Request — mit Diff, mit Review, und noch ein Jahr später " +
          "nachvollziehbar, wenn jemand fragt, wer die Formulierung zum " +
          "Zahlungsrückstand geändert hat. Tests halten fest, dass ein Dokument das " +
          "rendert, was Sie erwarten, sodass die CI den Fehler vor der Empfängerin " +
          "oder dem Empfänger bemerkt.",
      },
    },

    pullQuote:
      "Docxcelerate behandelt ein Dokument so, wie ein UI-Framework einen " +
      "Bildschirm behandelt: kleine Komponenten, zu einem Baum zusammengesetzt, " +
      "gerendert von etwas, das Papier versteht. Sie bekommen die Ergonomie eines " +
      "Komponentenmodells, und die Empfängerin oder der Empfänger bekommt ein " +
      "Word-Dokument.",

    engine: {
      title: "Die Engine",
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

    docs: {
      eyebrow: "Dokumentation",
      title: "Lesen Sie das Ganze",
      lead:
        "Jeder Node-Typ, jedes CLI-Flag und jede Datei, die ein Build schreibt — mit " +
        "Vorschauen, die der echte Renderer erzeugt hat, sodass nichts auf der Seite " +
        "einen Helper beschreiben kann, den es nicht mehr gibt.",
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
        staticAndDynamic: {
          title: "Statisch und dynamisch",
          blurb: "Was lokal läuft und was die Engine braucht.",
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
      "Eine echte Dokumentvorschau, live gerendert — keine Screenshots und keine versteckten Tricks.",
    counts: "{documents} Dokumente · {files} Dateien",
    sectors: {
      education: "Bildung",
      housing: "Wohnen",
      insurance: "Versicherung",
    },
    documents: {
      "offer-of-admission": "Zulassungsbescheid",
      "repairs-appointment": "Reparaturtermin",
      "policy-renewal": "Vertragsverlängerung",
    },
  },

  docs: {
    sidebarLabel: "Dokumentation",
    tocLabel: "Auf dieser Seite",
    onThisPage: "Auf dieser Seite",
    editOnGithub: "Diese Seite auf GitHub bearbeiten ↗",
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
    noHelpers: "Keine — von Hand geschrieben",
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
