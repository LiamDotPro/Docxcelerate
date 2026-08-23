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
    themes: "Thema's",
    components: "Componenten",
    cloud: "Cloud",
    github: "GitHub",
    switchTheme: "Thema wisselen",
    language: "Taal",
  },

  footer: {
    licence:
      "Documenten als componenten, DOCX als uitvoer. In de open ontwikkeld, onder " +
      "MIT-licentie, en draaiend op je eigen machine tenzij je anders besluit.",
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
    copyInstall: "Installatiecommando kopiëren",
    copy: "Kopiëren",
    copied: "Gekopieerd",
    copyFallback: "Druk op ⌘C",
    integrityTitle: "{hash} · klik om te kopiëren",
    /** Accessible name for the registry marks beside the version. */
    onNpm: "Docxcelerate op npm",
    onJsr: "Docxcelerate op JSR",

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
      ai: {
        label: "AI",
        title: "AI op componentniveau",
        body:
          "AI komt binnen via hooks, in de componenten die je toch al schrijft. Een " +
          "component geeft het model zijn context en wat je geschreven wilt hebben, " +
          "zodat het precies dat ene deel van het document maakt terwijl alles " +
          "eromheen deterministisch blijft. Jij bepaalt hoeveel er gegenereerd " +
          "wordt, één component tegelijk.",
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

    scale: {
      eyebrow: "Enterprise",
      title: "Gebouwd voor de rapporten die je nu met de hand schrijft",
      lead:
        "Een kantoor dat hetzelfde rapport honderden keren per maand maakt, heeft de " +
        "opmaak en het meeste van de tekst allang. Wat verandert, is de persoon voor " +
        "wie het geschreven wordt. Bouw dat rapport opnieuw op uit componenten, houd " +
        "elke run identiek, en geef een model alleen de delen die afhangen van wie " +
        "het leest.",
      traits: {
        volume: {
          title: "Eén template, elke ontvanger",
          body:
            "Publiceer de template één keer en roep hem daarna per persoon aan. Eén " +
            "rapport of honderdduizend is dezelfde aanroep, herhaald.",
        },
        determinism: {
          title: "Elke keer hetzelfde document",
          body:
            "Alles wat je niet als gegenereerd hebt gemarkeerd, rendert bij elke run " +
            "identiek. Alleen de delen die jij koos mogen variëren.",
        },
        integration: {
          title: "Aangeroepen door de systemen die je draait",
          body:
            "Je CRM, zaaksysteem of facturatiesysteem stuurt zijn gegevens en krijgt " +
            "een .docx terug. Niemand exporteert een spreadsheet of opent Word.",
        },
        reproducibility: {
          title: "Een jaar later nog reproduceerbaar",
          body:
            "Templates hebben een versie, dus elk document kan opnieuw worden " +
            "gebouwd uit dezelfde template en dezelfde gegevens. Een audit krijgt een " +
            "build, geen archief.",
        },
      },
    },

    openSource: {
      eyebrow: "Open source",
      title: "Open source, en gebouwd om dat te blijven",
      bodyOne:
        "Het framework, de renderers, het nodemodel en de CLI staan onder een " +
        "MIT-licentie en worden in de openbaarheid ontwikkeld. Lees de code die je " +
        "documenten schrijft, fork hem, of neem hem op in je eigen build.",
      bodyTwo:
        "De engine mag je gratis zelf hosten, dus documenten op schaal maken hangt " +
        "nooit af van een leverancier die blijft bestaan of van een prijslijst die " +
        "gelijk blijft. Onze betaalde cloud legt de premiumfuncties op diezelfde " +
        "vrije kern, zodat hosting en schaal er zijn vanaf het eerste document dat " +
        "je schrijft. Hij is het gemak, niet de deur naar binnen.",
      ctaSource: "Bekijk de broncode →",
      ctaNpm: "Bekijk het op npm",

      licence: {
        pages: "1 van 1",
        copyright: "Copyright (c) {year} Docxcelerate",
        fork: "Fork hem",
        vendor: "Neem hem op",
        ship: "Lever hem uit",
        read: "Lees hem →",
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

    agentSkill: {
      eyebrow: "Agent skills",
      title: "Geef het geheel aan je agent",
      lead:
        "Één Markdown-bestand leert een coding agent hoe documenten hier in elkaar " +
        "zitten: het componentmodel, de regels waar agents over struikelen en elk " +
        "commando. Zet het erin en vraag om een document in plaats van het eerste " +
        "zelf te schrijven.",
      where: "Waar het bestand hoort",
      copySkill: "Kopieer de skill",
      references:
        "Het is gewone Markdown, en er staan vier referentiebestanden naast in " +
        "skills/docxcelerate/.",
      agents: {
        "claude-code":
          "Zet de map in .claude/skills/ voor één project, of in ~/.claude/skills/ " +
          "voor allemaal. Hij laadt zichzelf zodra er een documentproject opduikt.",
        "cursor":
          "Sla het op als rule-bestand, dan leest Cursor het in dat project. Of " +
          "@-mention het bestand in de chat wanneer je het nodig hebt.",
        "copilot":
          "Plak het in .github/copilot-instructions.md, dan past Copilot het toe in " +
          "de hele repository.",
        "agents-md":
          "Codex, Gemini CLI, Aider en Cline lezen allemaal AGENTS.md in de root van " +
          "de repo. Plak het erin, of verwijs ernaar als je agent bestanden opent.",
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
      consulting: "Zakelijke dienstverlening",
    },
    documents: {
      "offer-of-admission": "Toelatingsbrief",
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
    reads: "Wat het leest",
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
      heading: "Begin met iets dat al werkt",
      body:
        "Thema's bepalen letter, kleur en pagina; componenten zijn nodes die " +
        "al doordacht zijn. Eén commando zet ze in een project, als bestanden " +
        "die van jou zijn.",
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
