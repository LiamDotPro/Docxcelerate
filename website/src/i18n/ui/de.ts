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

    scale: {
      eyebrow: "Im großen Maßstab",
      title: "Gebaut für ganze Dokumentbestände, nicht für Einzelstücke",
      lead:
        "Große Mengen waren eine Grundannahme von Anfang an, nichts nachträglich " +
        "Angebautes. Das Framework ist auf Organisationen zugeschnitten, die " +
        "Schriftverkehr zu Hunderttausenden versenden: dasselbe Zulassungsschreiben " +
        "an jede Bewerberin und jeden Bewerber, dieselbe Verlängerung an jede " +
        "Versicherungsnehmerin und jeden Versicherungsnehmer, jedes Mal mit anderen " +
        "Daten und anderen Klauseln.",
      traits: {
        volume: {
          title: "Ein Template, beliebig viele Dokumente",
          body:
            "Ein Template wird einmal veröffentlicht und dann pro Datensatz " +
            "aufgerufen. Hunderttausend Dokumente zu erzeugen ist derselbe Vorgang " +
            "wie eines zu erzeugen, nur wiederholt — Maßstab heißt hier, wie oft Sie " +
            "die Engine aufrufen, und nicht eine eigene Batch-Pipeline, die jemand " +
            "bauen und danach betreuen muss.",
        },
        determinism: {
          title: "Vorhersagbar von Haus aus",
          body:
            "Alles außer den Nodes, die Sie ausdrücklich als generiert markieren, " +
            "rendert jedes Mal identisch. Über einen ganzen Kundenbestand hinweg " +
            "zählt das mehr als alles andere: Sie können nachvollziehen, was jede " +
            "Empfängerin und jeder Empfänger bekommt, weil nur die Teile variieren " +
            "dürfen, die Sie ausgewählt haben.",
        },
        integration: {
          title: "Aufgerufen von den Systemen, die Sie schon betreiben",
          body:
            "Die Engine nimmt Daten über HTTP entgegen und gibt eine .docx zurück. " +
            "Ihr CRM, Ihre Fallverwaltung oder Ihr Abrechnungssystem erzeugt seinen " +
            "Schriftverkehr selbst — ohne dass jemand eine Tabelle exportiert, Word " +
            "öffnet oder ein Serienbrief-Makro pflegt, das nur eine Person versteht.",
        },
        reproducibility: {
          title: "Noch ein Jahr später reproduzierbar",
          body:
            "Ein veröffentlichtes Template ist versioniert, sodass sich das Dokument, " +
            "das eine Empfängerin oder ein Empfänger erhalten hat, exakt erneut " +
            "erzeugen lässt: aus demselben Template und denselben Daten. Wenn eine " +
            "Beschwerde oder eine Prüfung fragt, was im März verschickt wurde, ist " +
            "die Antwort ein Build und kein Archiv, von dem Sie hoffen, dass es " +
            "jemand aufbewahrt hat.",
        },
      },
    },

    openSource: {
      eyebrow: "Open Source",
      title: "Open Source, und darauf angelegt, es zu bleiben",
      bodyOne:
        "Das Framework, die Renderer, das Node-Modell und die CLI sind MIT-lizenziert " +
        "und werden offen entwickelt. Lesen Sie den Code, der Ihre Dokumente " +
        "schreibt, forken Sie ihn oder nehmen Sie ihn in Ihren eigenen Build auf. " +
        "Schreiben, Vorschau und das Packen von DOCX laufen vollständig auf Ihrem " +
        "Rechner — kein Konto, kein Upload, kein Netzwerkaufruf.",
      bodyTwo:
        "Die Generierungs-Engine lässt sich kostenlos selbst hosten. Dokumente im " +
        "großen Maßstab zu erzeugen hängt damit nie daran, dass ein Anbieter am " +
        "Markt bleibt oder eine Preisliste gleich bleibt. Die verwaltete Cloud ist " +
        "eine Bequemlichkeit für Teams, die sie lieber nicht selbst betreiben, und " +
        "nicht der einzige Weg hinein.",
      facts: {
        licence: {
          title: "MIT-Lizenz",
          body:
            "Kommerziell nutzen, verändern, ausliefern. Keine Nutzerzahlen, keine " +
            "Gebühr pro Dokument.",
        },
        local: {
          title: "Läuft auf Ihrem Rechner",
          body:
            "Dokumente und Daten bleiben auf Ihrem Laptop und in Ihrer CI, solange " +
            "Sie sie nicht an eine Engine senden, die Sie ausgewählt haben.",
        },
        selfHost: {
          title: "Selbst hostbare Engine",
          body:
            "Die kostenlose Engine läuft auf Ihrer eigenen Infrastruktur, in Ihrem " +
            "eigenen Netz.",
        },
      },
      ctaSource: "Zum Quellcode →",
      ctaNpm: "Auf npm ansehen",
    },

    docs: {
      eyebrow: "Dokumentation",
      title: "Lesen Sie das Ganze",
      lead:
        "Jeder Node-Typ, jedes CLI-Flag und jede Datei, die ein Build schreibt. Dazu " +
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
