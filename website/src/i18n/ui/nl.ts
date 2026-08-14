/**
 * Dutch UI strings.
 *
 * "node", "engine", "template", "workspace", "preview" and "helper" stay
 * English: they are the words the API uses and the words Dutch developers use
 * for them. Translating them would put a name on the page that appears nowhere
 * in the code the page is describing.
 */
import type { UiStrings } from "./en";

export const nl: UiStrings = {
  meta: {
    tagline: "Documenten als componenten. DOCX als resultaat.",
    description:
      "Een TypeScript-toolkit en CLI om DOCX-documentprojecten te bouwen. Schrijf " +
      "documenten als getypeerde componenten, bekijk ze in de browser en pak echte " +
      ".docx-bestanden lokaal in.",
  },

  common: {
    skipToContent: "Naar de inhoud",
    soon: "binnenkort",
  },

  nav: {
    label: "Hoofdnavigatie",
    docs: "Docs",
    cloud: "Cloud",
    github: "GitHub",
    switchTheme: "Thema wisselen",
    language: "Taal",
  },

  footer: {
    label: "Voettekst",
    licence:
      "MIT-licentie. Schrijven, preview en het inpakken van DOCX draaien volledig op je eigen machine.",
    docs: "Docs",
    cloud: "Cloud",
    github: "GitHub",
    npm: "npm",
    copyright: "© {year} Docxcelerate-bijdragers.",
  },

  home: {
    headlineLineOne: "Documenten als componenten.",
    headlineLineTwo: "DOCX als resultaat.",
    lead:
      "Stel documenten samen uit kleine, getypeerde componenten, met de JSX die je " +
      "toch al schrijft. Gebruik de ingebouwde AI-functies om dynamische documenten " +
      "te maken en schaal documentgeneratie op met onze engine.",
    ctaStart: "Aan de slag →",
    ctaHowItWorks: "Hoe het werkt",
    copyInstall: "Installatiecommando kopiëren",
    copy: "Kopiëren",
    copied: "Gekopieerd",
    copyFallback: "Druk op ⌘C",
    integrityTitle: "{hash} — klik om te kopiëren",
    /** Accessible name for the registry marks beside the version. */
    onNpm: "Docxcelerate op npm",
    onJsr: "Docxcelerate op JSR",

    points: {
      authoring: {
        label: "Schrijven",
        title: "Schrijf documenten zoals websites",
        body:
          "Een document is een boom van getypeerde componenten. Als je ooit React " +
          "hebt geschreven, ken je deze vorm al — props, compositie, kleine " +
          "bestanden — dus een frontend-engineer is de eerste middag al productief " +
          "in plaats van eerst een templatetaal te moeten leren.",
      },
      ai: {
        label: "AI",
        title: "Ontworpen voor AI",
        body:
          "Gegenereerde tekst is een nodetype, geen aanbouwsel. Markeer de ene " +
          "alinea die echt een model nodig heeft, geef die prompts en een " +
          "placeholder, en laat de rest deterministisch — zo lost AI het moeilijke " +
          "deel op zonder dat het hele document aan zijn genade is overgeleverd.",
      },
      changeControl: {
        label: "Wijzigingsbeheer",
        title: "Documenten staan in je repository",
        body:
          "Omdat een document broncode is, is het wijzigen van een zin een pull " +
          "request — gedift, beoordeeld, en een jaar later nog steeds herleidbaar " +
          "wanneer iemand vraagt wie de formulering over betalingsachterstanden " +
          "heeft aangepast. Tests leggen vast dat een document rendert wat je " +
          "verwacht, zodat CI de fout opmerkt vóór de ontvanger dat doet.",
      },
    },

    pullQuote:
      "Docxcelerate behandelt een document zoals een UI-framework een scherm " +
      "behandelt: kleine componenten, samengevoegd tot een boom, gerenderd door " +
      "iets dat verstand heeft van papier. Jij krijgt het gemak van een " +
      "componentmodel, en de ontvanger krijgt een Word-document.",

    engine: {
      title: "De engine",
      bodyOne:
        "In de engine worden documenten daadwerkelijk geschreven. Hij vult je " +
        "gegevens in, voert de AI uit en geeft het voltooide document terug. Elk " +
        "soort node kan AI gebruiken, niet alleen alinea's. Het antwoord van het " +
        "model wordt ofwel de tekst, geschreven op basis van de informatie die je " +
        "meegeeft, ofwel een beslissing waarvan het document afhangt.",
      bodyTwo:
        "Je publiceert een template één keer naar de engine. Daarna kan elk " +
        "systeem zijn API aanroepen met een set gegevens en krijgt het een " +
        "document terug. Er is een gratis engine die je zelf kunt hosten. De " +
        "managed cloud draait de volledige versie, met veel dat de gratis versie " +
        "niet heeft, en komt binnenkort.",
      ctaHow: "Hoe de engine werkt →",
      ctaCloud: "Managed cloud",
      steps: {
        build: {
          title: "Bouwen",
          detail:
            "Het framework maakt van je document een pakket. Deze stap draait op je eigen machine.",
        },
        publish: {
          title: "Publiceren",
          detail:
            "Je stuurt het pakket naar een engine. De engine slaat het op en geeft het een naam.",
        },
        write: {
          title: "Schrijven",
          detail:
            "Je applicatie stuurt een set gegevens. De engine geeft het voltooide document terug.",
        },
      },
    },

    docs: {
      eyebrow: "Documentatie",
      title: "Lees het helemaal",
      lead:
        "Elk nodetype, elke CLI-vlag en elk bestand dat een build wegschrijft — met " +
        "previews die door de echte renderer zijn gerenderd, zodat niets op de " +
        "pagina een helper kan beschrijven die niet meer bestaat.",
      cards: {
        startHere: {
          title: "Begin hier",
          blurb: "Zet een workspace op en open je eerste document.",
        },
        documentsAndNodes: {
          title: "Documenten en nodes",
          blurb: "Het componentmodel waaruit een document is opgebouwd.",
        },
        nodeModel: {
          title: "Het nodemodel",
          blurb: "Elk nodetype, met van elk een preview.",
        },
        staticAndDynamic: {
          title: "Statisch en dynamisch",
          blurb: "Wat lokaal draait, en wat de engine nodig heeft.",
        },
        cli: {
          title: "CLI-commando's",
          blurb: "dxcl init, new, node.",
        },
        entrypoints: {
          title: "Package-entrypoints",
          blurb: "Alles wat importeerbaar is, en wat het exporteert.",
        },
      },
    },
  },

  demo: {
    showFiles: "Workspacebestanden tonen",
    hideFiles: "Workspacebestanden verbergen",
    workspaceFiles: "Workspacebestanden",
    resize: "Workspace en preview verdelen",
    resizeHint: "Sleep om te verdelen · dubbelklik om te herstellen",
    zoom: "Preview zoomen",
    zoomIn: "Inzoomen",
    zoomOut: "Uitzoomen",
    fit: "Passend",
    frameTitle: "Het geselecteerde document, gerenderd door Docxcelerate",
    caption:
      "Een echte documentpreview die live wordt gerenderd, geen screenshots en geen verborgen trucs.",
    counts: "{documents} documenten · {files} bestanden",
    sectors: {
      education: "Onderwijs",
      housing: "Wonen",
      insurance: "Verzekeringen",
    },
    documents: {
      "offer-of-admission": "Toelatingsbrief",
      "repairs-appointment": "Reparatieafspraak",
      "policy-renewal": "Polisverlenging",
    },
  },

  docs: {
    sidebarLabel: "Documentatie",
    tocLabel: "Op deze pagina",
    onThisPage: "Op deze pagina",
    editOnGithub: "Deze pagina bewerken op GitHub ↗",
    untranslatedNote:
      "Deze pagina is nog niet vertaald en wordt daarom in het Engels getoond.",

    groups: {
      "Start Here": "Begin hier",
      "Essentials": "Basis",
      "Nodes": "Nodes",
      "CLI": "CLI",
      "Projects": "Projecten",
      "Generation": "Generatie",
      "Reference": "Referentie",
    },
    subgroups: {
      "Node types": "Nodetypes",
    },
  },

  nodes: {
    helpers: "Helpers",
    noHelpers: "Geen — met de hand geschreven",
    kind: "Nodesoort",
    category: "Categorie",
    resolves: "Wordt opgelost",
    children: "Children",
    option: "Optie",
    type: "Type",
    whatItDoes: "Wat het doet",
    required: "verplicht",
    renderNote: "Wat de renderers vandaag doen",
    resolvesTo: "Waartoe het wordt opgelost",
    resolvesToBody:
      "De node zoals hij in het {model} verschijnt: de JSON die een renderer " +
      "aangereikt krijgt. Geen opmaak, geen lay-out.",
    endpointAsked: "Wat er aan het endpoint wordt gevraagd",
    endpointAskedBody:
      "Opgelost tegen dezelfde voorbeelddata. Een previewbuild stopt bij de " +
      "placeholder; een build op het moment van de aanvraag stuurt deze mee.",
    previewTitle: "{type}: {variant}, gerenderd",
    reference: "{title}-referentie, met previews →",
    categories: {
      Structure: "Structuur",
      Text: "Tekst",
      Media: "Media",
      Data: "Data",
    },
    status: {
      "stable": "Stabiel",
      "no-helper": "Nog geen helper",
      "planned": "Gepland",
    },
  },

  notFound: {
    title: "Pagina niet gevonden",
    description: "Die pagina bestaat niet.",
    heading: "Die pagina bestaat niet",
    body:
      "De link is misschien verouderd, of de pagina is verplaatst. De documentatie " +
      "is de beste plek om de draad weer op te pakken.",
    ctaDocs: "Lees de documentatie →",
    ctaHome: "Terug naar de homepage",
  },
};
