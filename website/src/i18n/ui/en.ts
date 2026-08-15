/**
 * English UI strings — the source of truth.
 *
 * Every other locale is typed against this object, so adding a key here is a
 * type error in the four translations until they carry it too. Keys are named
 * for what the string *is*, never for the English words in it.
 *
 * What is deliberately not here: anything generated from the toolkit source.
 * The node catalog's summaries, option descriptions and render notes are built
 * by scripts/build-node-previews.mjs out of src/nodes/, so a translation of
 * them would go stale the moment the toolkit changed, silently. Labels *around*
 * that content are translated; the content itself stays English.
 */
export const en = {
  /**
   * What search engines and link previews see. The product name itself is not
   * here — Docxcelerate is spelled the same in every language.
   */
  meta: {
    tagline: "Documents as components. DOCX as output.",
    description:
      "A TypeScript toolkit and CLI for building DOCX document projects. Author " +
      "documents as typed components, preview them in the browser, and pack real " +
      ".docx files locally.",
  },

  common: {
    skipToContent: "Skip to content",
    soon: "soon",
  },

  nav: {
    label: "Main",
    docs: "Docs",
    cloud: "Cloud",
    github: "GitHub",
    switchTheme: "Switch theme",
    /** Names the control itself, not the language it is currently set to. */
    language: "Language",
  },

  footer: {
    label: "Footer",
    licence:
      "MIT licensed. Authoring, preview and DOCX packing run entirely on your machine.",
    docs: "Docs",
    cloud: "Cloud",
    github: "GitHub",
    npm: "npm",
    /** {year} is substituted at render time. */
    copyright: "© {year} Docxcelerate contributors.",
  },

  home: {
    /** Two lines, broken deliberately — the break is part of the composition. */
    headlineLineOne: "Documents as components.",
    headlineLineTwo: "DOCX as output.",
    lead:
      "Compose documents from small typed components, using the JSX you already " +
      "write. Built-in AI can write any part of a document, or decide what goes " +
      "in it. Then generate them at scale through the engine.",
    ctaStart: "Get started →",
    ctaHowItWorks: "How it works",
    copyInstall: "Copy install command",
    copy: "Copy",
    copied: "Copied",
    copyFallback: "Press ⌘C",
    /** {hash} is the full integrity hash. */
    integrityTitle: "{hash} · click to copy",
    /** Accessible name for the registry marks beside the version. */
    onNpm: "Docxcelerate on npm",
    onJsr: "Docxcelerate on JSR",

    points: {
      authoring: {
        label: "Authoring",
        title: "Write documents like websites",
        body:
          "A document is a tree of typed components. If you have written React you " +
          "already know the shape of this: props, composition, small files. A " +
          "frontend engineer is productive on the first afternoon rather than " +
          "learning a template language first.",
      },
      ai: {
        label: "AI",
        title: "Designed for AI",
        body:
          "Generated prose is a node type, not a bolt-on. Mark the one paragraph " +
          "that genuinely needs a model, give it prompts and a placeholder, and " +
          "leave the rest deterministic. AI solves the hard part without putting " +
          "the whole document at its mercy.",
      },
      changeControl: {
        label: "Change control",
        title: "Documents live in your repo",
        body:
          "Because a document is source, changing a sentence is a pull request that " +
          "gets diffed, reviewed, and stays attributable a year later when someone " +
          "asks who altered the arrears wording. Tests assert a document renders what you " +
          "expect, so CI catches the mistake before a recipient does.",
      },
    },

    pullQuote:
      "Docxcelerate treats a document the way a UI framework treats a screen: " +
      "small components, composed into a tree, rendered by something that " +
      "knows about paper. You get the ergonomics of a component model, and " +
      "the recipient gets a Word document.",

    engine: {
      title: "The engine",
      bodyOne:
        "The engine is where documents are actually written. It fills in your " +
        "data, runs the AI, and returns the finished document. Any kind of node " +
        "can use AI, not only paragraphs. The model's answer either becomes the " +
        "text, written from the information you give it, or makes a decision " +
        "the document depends on.",
      bodyTwo:
        "You publish a template to the engine once. After that, any system can " +
        "call its API with a set of data and get a document back. A free engine " +
        "is available to self-host. The managed cloud runs the complete one, " +
        "with a lot the free version does not have, and is coming soon.",
      ctaHow: "How the engine works →",
      ctaCloud: "Managed cloud",
      steps: {
        build: {
          title: "Build",
          detail:
            "The framework turns your document into a package. This step runs on your machine.",
        },
        publish: {
          title: "Publish",
          detail: "You send the package to an engine. The engine stores it and gives it a name.",
        },
        write: {
          title: "Write",
          detail: "Your application sends a set of data. The engine returns the finished document.",
        },
      },
    },

    scale: {
      eyebrow: "At scale",
      title: "Built for document estates, not one-off files",
      lead:
        "Volume was a founding assumption, not something bolted on later. The " +
        "framework was shaped around organisations that send correspondence by the " +
        "hundred thousand: the same offer letter to every applicant, the same " +
        "renewal schedule to every policyholder, each one carrying different data " +
        "and different clauses.",
      traits: {
        volume: {
          title: "One template, any number of documents",
          body:
            "A template is published once and then called per record. Generating a " +
            "hundred thousand documents is the same operation as generating one, " +
            "repeated — scale is how often you call the engine, not a separate " +
            "batch pipeline somebody has to build and then babysit.",
        },
        determinism: {
          title: "Predictable by default",
          body:
            "Everything except the nodes you explicitly mark as generated renders " +
            "identically every time. Across a whole customer base that matters more " +
            "than anything else: you can reason about what every recipient will get, " +
            "because only the parts you chose are allowed to vary.",
        },
        integration: {
          title: "Called by the systems you already run",
          body:
            "The engine takes data over HTTP and returns a .docx. Your CRM, case " +
            "management or billing system generates its own correspondence, with " +
            "nobody exporting a spreadsheet, opening Word, or maintaining a " +
            "mail-merge macro that one person understands.",
        },
        reproducibility: {
          title: "Reproducible a year later",
          body:
            "A published template is versioned, so the document a recipient received " +
            "can be produced again exactly, from the same template and the same " +
            "data. When a complaint or an audit asks what was sent in March, the " +
            "answer is a build rather than an archive you hope someone kept.",
        },
      },
    },

    openSource: {
      eyebrow: "Open source",
      title: "Open source, and built to stay that way",
      bodyOne:
        "The framework, the renderers, the node model and the CLI are MIT licensed " +
        "and developed in the open. Read the code that writes your documents, fork " +
        "it, or vendor it into your own build. Authoring, preview and DOCX packing " +
        "run entirely on your machine — no account, no upload, no network call.",
      bodyTwo:
        "The generation engine is free to self-host, so running documents at scale " +
        "never depends on a vendor staying in business or on a price list staying " +
        "the same. The managed cloud is a convenience for teams who would rather " +
        "not operate it themselves, not the only door in.",
      facts: {
        licence: {
          title: "MIT licence",
          body: "Use it commercially, change it, ship it. No seat count, no per-document fee.",
        },
        local: {
          title: "Runs on your machine",
          body:
            "Documents and data stay on your laptop and in your CI unless you send " +
            "them to an engine you chose.",
        },
        selfHost: {
          title: "Self-hostable engine",
          body: "The free engine runs on your own infrastructure, inside your own network.",
        },
      },
      ctaSource: "View the source →",
      ctaNpm: "See it on npm",
    },

    docs: {
      eyebrow: "Documentation",
      title: "Read the whole thing",
      lead:
        "Every node type, every CLI flag, and every file a build writes, with " +
        "previews rendered by the real renderer, so nothing on the page can " +
        "describe a helper that no longer exists.",
      cards: {
        startHere: {
          title: "Start here",
          blurb: "Scaffold a workspace and open your first document.",
        },
        documentsAndNodes: {
          title: "Documents and nodes",
          blurb: "The component model a document is built from.",
        },
        nodeModel: {
          title: "The node model",
          blurb: "Every node type, with a preview of each.",
        },
        writingNodes: {
          title: "Writing nodes",
          blurb: "The pieces a document is made of.",
        },
        cli: {
          title: "CLI commands",
          blurb: "dxcl init, new, node.",
        },
        entrypoints: {
          title: "Package entrypoints",
          blurb: "Everything importable, and what it exports.",
        },
      },
    },
  },

  demo: {
    showFiles: "Show workspace files",
    hideFiles: "Hide workspace files",
    workspaceFiles: "Workspace files",
    resize: "Resize workspace and preview",
    resizeHint: "Drag to resize · double-click to reset",
    zoom: "Preview zoom",
    zoomIn: "Zoom in",
    zoomOut: "Zoom out",
    fit: "Fit",
    frameTitle: "The selected document, rendered by Docxcelerate",
    caption: "Real document preview being rendered, no screenshots and no hidden tricks.",
    /** {documents} and {files} are counts. */
    counts: "{documents} documents · {files} files",
    sectors: {
      education: "Education",
      housing: "Housing",
      insurance: "Insurance",
    },
    documents: {
      "offer-of-admission": "Offer of Admission",
      "repairs-appointment": "Repair Appointment",
      "policy-renewal": "Policy Renewal",
    },
  },

  docs: {
    sidebarLabel: "Documentation",
    tocLabel: "On this page",
    onThisPage: "On this page",
    editOnGithub: "Edit this page on GitHub ↗",
    /**
     * Shown when a page has no translation yet and the English original is
     * standing in for it. Better than a 404, and better than pretending: the
     * reader is told why the language changed under them.
     */
    untranslatedNote:
      "This page has not been translated into your language yet, so it is shown in English.",

    /**
     * Display names for the sidebar groups. The frontmatter values stay
     * English — they are keys, matched by the content schema — and this is
     * what a reader sees.
     */
    groups: {
      "Start Here": "Start Here",
      "Essentials": "Essentials",
      "Nodes": "Nodes",
      "CLI": "CLI",
      "Projects": "Projects",
      "Generation": "Generation",
      "Reference": "Reference",
    },
    subgroups: {
      "Node types": "Node types",
    },
  },

  nodes: {
    helpers: "Helpers",
    noHelpers: "None, written by hand",
    kind: "Node kind",
    category: "Category",
    resolves: "Resolves",
    children: "Children",
    option: "Option",
    type: "Type",
    whatItDoes: "What it does",
    required: "required",
    renderNote: "What the renderers do today",
    resolvesTo: "What it resolves to",
    /** {model} is rendered as a code tag around the DocumentModel type name. */
    resolvesToBody:
      "The node as it appears in the {model}: the JSON a renderer is handed. " +
      "No styling, no layout.",
    endpointAsked: "What the endpoint is asked",
    endpointAskedBody:
      "Resolved against the same sample data. A preview build stops at the " +
      "placeholder; a request-time build sends these.",
    /** {type} is the node type, {variant} the variant — both from the catalog. */
    previewTitle: "{type}: {variant}, rendered",
    /** {title} is the node type's name, e.g. "Paragraph". */
    reference: "{title} reference, with previews →",
    categories: {
      Structure: "Structure",
      Text: "Text",
      Media: "Media",
      Data: "Data",
    },
    status: {
      "stable": "Stable",
      "no-helper": "No helper yet",
      "planned": "Planned",
    },
  },

  notFound: {
    title: "Page not found",
    description: "That page does not exist.",
    heading: "That page does not exist",
    body:
      "The link may be out of date, or the page may have moved. The documentation " +
      "is the best place to pick the thread back up.",
    ctaDocs: "Read the docs →",
    ctaHome: "Back to the homepage",
  },
} as const;

/**
 * The same shape, with every leaf widened from its English literal to `string`.
 * Recursive, so it does not care how deeply a section nests.
 */
type Translated<T> = {
  readonly [Key in keyof T]: T[Key] extends string ? string : Translated<T[Key]>;
};

/**
 * The shape every locale must satisfy. Derived from English rather than
 * declared separately, so the two cannot disagree: a key added here is a type
 * error in all four translations until they carry it too.
 */
export type UiStrings = Translated<typeof en>;
