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
    registry: "Registry",
    cloud: "Cloud",
    github: "GitHub",
    switchTheme: "Thema wisselen",
    language: "Taal",
  },

  footer: {
    licence:
      "Documenten als componenten, DOCX als uitvoer. In de open ontwikkeld, onder " +
      "MIT-licentie, en draaiend op je eigen machine tenzij je anders besluit.",
    licenceFile: "De MIT-licentie lezen",
    docs: "Docs",
    project: "Project",
    engine: "Engine",
    cloud: "Cloud",
    github: "GitHub",
    npm: "npm",
    releases: "Releases",
    artifacts: "Build-artefacten",
    copyright: "© {year} Docxcelerate-bijdragers.",
  },

  home: {
    headlineLineOne: "Documenten als componenten.",
    headlineLineTwo: "DOCX als resultaat.",
    lead:
      "Stel documenten samen uit kleine, getypeerde componenten, met de JSX die je " +
      "toch al schrijft. Bouw moeiteloos AI in die de inhoud schrijft of erover " +
      "beslist. Benut onze engine om documenten op schaal te genereren.",
    ctaStart: "Aan de slag →",
    ctaHowItWorks: "Hoe het werkt",
    copied: "Gekopieerd",
    copyFallback: "Druk op ⌘C",
    integrityTitle: "{hash} · klik om te kopiëren",
    /** Accessible name for the registry marks beside the version. */
    onNpm: "Docxcelerate op npm",
    onJsr: "Docxcelerate op JSR",

    /** The hero's switch between the two builds of the document it shows. */
    builds: {
      label: "Eén template · twee builds",
      engine: "Engine",
      preview: "Preview",
      engineCaption: "Engine: de gemarkeerde delen zijn door een model voor deze klant geschreven, en de beslissing was aan het model. Elk bedrag wordt berekend, nooit gegenereerd.",
      previewCaption: "Preview: hetzelfde template op je eigen machine. Gegenereerde delen tonen hun placeholder, er wordt niets verstuurd en de pagina ziet er na elke save hetzelfde uit.",
    },

    /** §1: one paragraph followed from source to preview to the engine. */
    source: {
      title: "Bron. Preview. Document.",
      lead: "Eén alinea uit het portefeuilleoverzicht, gevolgd van het bestand dat je schrijft tot de pagina die je klant opent. De preview wacht nooit op een model en er verlaat niets je machine; de engine schrijft de echte zin per persoon.",
      sourceLabel: "bron",
      previewLabel: "preview",
      sourceNote: "Eén keer geschreven, in je repo. De prompt staat naast de node die hij vult.",
      previewNote: "Opnieuw gebouwd bij elke save. Gegenereerde nodes tonen hun placeholder.",
      engineNote: "Per ontvanger geschreven. De zinnen eromheen veranderen nooit.",
    },

    points: {
      authoring: {
        label: "Schrijven",
        title: "Schrijf documenten zoals websites",
        body:
          "Een document is een boom van getypeerde componenten. Als je ooit React " +
          "hebt geschreven, ken je deze vorm al: props, compositie, kleine " +
          "bestanden. Een frontend-engineer is de eerste middag al productief " +
          "in plaats van eerst een templatetaal te moeten leren.",
      },
      changeControl: {
        label: "Wijzigingsbeheer",
        title: "Documenten staan in je repository",
        body:
          "Omdat een document broncode is, is het wijzigen van een zin een pull " +
          "request: gedift, beoordeeld, en een jaar later nog steeds herleidbaar " +
          "wanneer iemand vraagt wie de formulering over betalingsachterstanden " +
          "heeft aangepast. Tests leggen vast dat een document rendert wat je " +
          "verwacht, zodat CI de fout opmerkt vóór de ontvanger dat doet.",
      },
    },

    pullQuote:
      "Een document is een boom van kleine componenten, gerenderd door iets dat " +
      "verstand heeft van papier. Jij krijgt het gemak van een componentmodel. " +
      "De persoon aan de andere kant krijgt een Word-bestand.",

    attribution: {
      name: "Liam",
      role: "Auteur, Docxcelerate",
      portraitAlt: "Liam, de auteur van Docxcelerate",
    },

    engine: {
      eyebrow: "De engine",
      title: "Publiceer één keer. Schaal daarna op.",
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
      /** {count} is a number of documents. */
      volume: "{count} overzichten dit kwartaal",
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
      title: "Alles is gedocumenteerd",
      lead:
        "Elk nodetype, elke CLI-vlag en elk bestand dat een build wegschrijft, met " +
        "previews die door de echte renderer zijn gerenderd, zodat niets op de " +
        "pagina een helper kan beschrijven die niet meer bestaat.",
      ctaAll: "Volledige documentatie →",
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
        writingNodes: {
          title: "Nodes schrijven",
          blurb: "De onderdelen waaruit een document bestaat.",
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
      housing: "Wonen",
      insurance: "Verzekeringen",
      consulting: "Zakelijke dienstverlening",
      finance: "Financiën",
    },
    documents: {
      "portfolio-review": "Kwartaaloverzicht portefeuille",
      "repairs-appointment": "Reparatieafspraak",
      "policy-renewal": "Polisverlenging",
      invoice: "Factuur",
    },
  },

  docs: {
    sidebarLabel: "Documentatie",
    tocLabel: "Op deze pagina",
    onThisPage: "Op deze pagina",
    editOnGithub: "Deze pagina bewerken op GitHub ↗",
    viewAsMarkdown: "Deze pagina als Markdown lezen",
    /** {version} is the release the page was built against. */
    renderedAgainst: "Gerenderd met {version}",
    previous: "Vorige",
    next: "Volgende",
    /** Names the previous/next links at the foot of a page. */
    pagination: "Vorige en volgende pagina",
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
    noHelpers: "Geen, met de hand geschreven",
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

  registry: {
    /** The detail pages, per design boards 2b and 2c. */
    backToRegistry: "Registry",
    renderedBy: "Gerenderd door Docxcelerate tijdens de build",
    openPreview: "Openen ↗",
    specification: "Specificatie",
    specNote: "Waarden komen uit style.page en style.typography.",
    margins: "Marges",
    body: "Broodtekst",
    titleLabel: "Titel",
    sectionHeading: "Sectiekop",
    paragraphLabel: "Alinea",
    after: "erna",
    compare: "Vergelijken",
    compareNote: "Dezelfde alinea in elk thema",
    viewing: "in beeld",
    requires: "Vereist",
    requiresNothing: "Niets. Het installeert alleen.",
    installsAlone: "Het installeert alleen.",
    stepData: "Voorbeelddata",
    stepDataNote: "de JSON waarmee dit voorbeeld is gebouwd",
    stepPreview: "Gerenderd voorbeeld",
    stepPreviewNote: "de component op zichzelf gerenderd",
    stepPreviewArrangementsNote: "elke opstelling op zichzelf gerenderd",
    /** Marks the arrangement a component uses when none is asked for. */
    arrangementDefault: "standaard",
    /** Labels the picker that chooses which arrangement the preview shows. */
    arrangementLabel: "Opstelling",
    /** The picker's option for showing every arrangement at once. */
    arrangementAll: "Alle opstellingen",
    stepNodes: "Opgeloste nodes",
    /** {count} is a number of nodes. */
    stepNodesNote: "wat een renderer krijgt — {count} nodes",
    stepFiles: "De bestanden",
    stepFilesNote: "wat installeren schrijft; het gekozen bestand is de component",
    dataFieldsNote: "Voeg deze na het installeren toe aan je types.ts. Dat is de enige handmatige stap.",
    copyFieldPaths: "Veldpaden kopiëren",
    dataFields: "Datavelden",
    copy: "Kopiëren",
    copied: "Gekopieerd",
    copyFile: "{file} kopiëren",
    /** {exports} is a list of names, {count} a number of nodes. */
    exportsResolves: "exporteert {exports} · levert {count} nodes op",
    categoryTags: "Categorie · tags",
    nothingHere: "Hier staat nog niets.",
    nothingHereNote: "De registry wordt uit het pakket gebouwd. Zodra een item meegaat, verschijnt het hier.",
    /** The one-page registry at /registry. */
    browse: {
      title: "Registry",
      description:
        "Kant-en-klare themas en documentcomponenten voor Docxcelerate, met één commando in een project geïnstalleerd.",
      eyebrow: "Registry",
      heading: "Registry",
      lead:
        "Kant-en-klare themas en componenten voor je documentprojecten. Installeer er een met één commando en bewerk het daarna als je eigen code.",
      allTags: "Alles",
      themesNote: "Een thema installeren schrijft document-style.ts",
      nodes: "nodes",
      noPreview: "Geen voorbeeld",
      /** Opent de datapaden op een componentrij; verbindt gefilterde tags. */
      readsPrefix: "leest",
      /** The page's headline; `heading` is the short name the nav uses. */
      headline: "Begin met papier dat al werkt.",
      /** {themes} and {components} are counts. */
      counts: "{themes} thema's · {components} componenten",
      useTheme: "Thema gebruiken",
      selected: "✓ Gekozen",
      add: "+ Toevoegen",
      added: "✓ Toegevoegd",
      familyLabel: "Wat het tekent",
      familyNote: "Gegroepeerd op waar een component in uitkomt, gelezen uit de build — geen label dat iemand bijhoudt.",
      /** {count} is how many themes and components are picked. */
      trayCount: "Je installatie · {count} gekozen",
      trayEmpty: "kies een thema of een component",
    },
    themes: {
      title: "Thema's",
      description:
        "Kant-en-klare documentthema's voor Docxcelerate: lettertypen, " +
        "kleuren, witruimte en paginaformaat, met één commando geïnstalleerd.",
      heading: "Thema's",
      lead:
        "Lettertypen, kleuren, witruimte en pagina, onder één naam. Een thema " +
        "is één import in een project en reist mee met het gebouwde document — " +
        "wie het rendert hoeft niets te installeren.",
      cta: "Bekijk de thema's",
    },
    components: {
      title: "Componenten",
      description:
        "Kant-en-klare Docxcelerate-componenten: briefhoofden, adresblokken, " +
        "betalingsoverzichten en meer, als broncode in je eigen project.",
      heading: "Componenten",
      lead:
        "Nodes die iemand al geschreven en doordacht heeft. Toevoegen kopieert " +
        "de broncode naar je project — geen versie, geen upgradepad, niets om " +
        "te forken.",
      cta: "Bekijk de componenten",
    },
    install: "Installeren",
    preview: "Voorbeeld",
    category: "Categorie",
    tags: "Labels",
    fonts: "Lettertypen",
    page: "Pagina",
    typography: "Typografie",
    colours: "Kleuren",
    spacing: "Witruimte",
    exports: "Exporteert",
    field: "Veld",
    type: "Type",
    whatItDoes: "Waar het voor is",
    source: "Broncode",
    lands: "Komt terecht in",
    resolvesTo: "Wat eruit komt",
    previewData: "Voorbeeld met",
    drawnAgainst: "Ontworpen op het thema {theme}.",
    allThemes: "Alle thema's",
    allComponents: "Alle componenten",
    ad: {
      heading: "Begin je volgende document met een voorsprong.",
      body:
        "De registry bevat themas en documentcomponenten. `dxcl add` kopieert het bestand naar je project. Geen dependency, geen versie om bij te houden.",
      browse: "Bekijk de registry",
      howItWorks: "Hoe installeren werkt",
      written: "2 bestanden geschreven · geen dependency toegevoegd",
      afterInstall:
        "Na het installeren staan de bestanden in je repo. Bewerk ze als elke andere broncode.",
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
