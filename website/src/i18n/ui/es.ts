/**
 * Spanish UI strings.
 *
 * Addresses the reader as "tú", which is what Spanish developer documentation
 * overwhelmingly uses and what matches the directness of the English.
 *
 * "nodo", "plantilla", "motor" and "vista previa" are translated — Spanish
 * technical writing translates all four and the English words would read as
 * untranslated rather than as terminology. "helper", "workspace", "prompt" and
 * "pull request" stay English: they name things that appear verbatim in the
 * API, the CLI output and the tooling around it.
 */
import type { UiStrings } from "./en";

export const es: UiStrings = {
  meta: {
    tagline: "Documentos como componentes. DOCX como resultado.",
    description:
      "Un toolkit de TypeScript y una CLI para crear proyectos de documentos DOCX. " +
      "Escribe documentos como componentes tipados, previsualízalos en el navegador y " +
      "empaqueta archivos .docx reales en local.",
  },

  common: {
    skipToContent: "Ir al contenido",
    soon: "pronto",
  },

  nav: {
    label: "Principal",
    docs: "Docs",
    registry: "Registro",
    cloud: "Cloud",
    github: "GitHub",
    switchTheme: "Cambiar de tema",
    language: "Idioma",
  },

  footer: {
    licence:
      "Documentos como componentes, DOCX como salida. Desarrollado en abierto, con " +
      "licencia MIT, y ejecutándose en tu propia máquina salvo que decidas otra cosa.",
    licenceFile: "Leer la licencia MIT",
    docs: "Docs",
    project: "Proyecto",
    engine: "Motor",
    cloud: "Cloud",
    github: "GitHub",
    npm: "npm",
    releases: "Versiones",
    artifacts: "Artefactos de compilación",
    copyright: "© {year} Colaboradores de Docxcelerate.",
  },

  home: {
    headlineLineOne: "Documentos como componentes.",
    headlineLineTwo: "DOCX como resultado.",
    lead:
      "Compón documentos a partir de componentes pequeños y tipados, con el JSX que " +
      "ya escribes. Incorpora IA sin esfuerzo para escribir el contenido o decidir " +
      "sobre él. Aprovecha nuestro motor para generar documentos a gran escala.",
    ctaStart: "Empezar →",
    ctaHowItWorks: "Cómo funciona",
    copied: "Copiado",
    copyFallback: "Pulsa ⌘C",
    integrityTitle: "{hash} · haz clic para copiar",
    /** Accessible name for the registry marks beside the version. */
    onNpm: "Docxcelerate en npm",
    onJsr: "Docxcelerate en JSR",

    /** The hero's switch between the two builds of the document it shows. */
    builds: {
      label: "Una plantilla · dos builds",
      engine: "Motor",
      preview: "Vista previa",
      engineCaption: "Motor: las partes resaltadas las escribió un modelo para este cliente, y la decisión fue suya. Cada cifra se calcula, nunca se genera.",
      previewCaption: "Vista previa: la misma plantilla en tu máquina. Las partes generadas muestran su texto provisional, no se envía nada a ningún sitio y la página sale igual cada vez que guardas.",
    },

    /** §1: one paragraph followed from source to preview to the engine. */
    source: {
      title: "Código. Vista previa. Documento.",
      lead: "Un párrafo del informe de cartera, seguido desde el archivo que escribes hasta la página que abre tu cliente. La vista previa nunca espera a un modelo y nada sale de tu máquina; el motor escribe la frase real para cada persona.",
      sourceLabel: "código",
      previewLabel: "vista previa",
      sourceNote: "Se escribe una vez, en tu repositorio. El prompt vive junto al nodo que rellena.",
      previewNote: "Se reconstruye cada vez que guardas. Los nodos generados muestran su texto provisional.",
      engineNote: "Se escribe para cada destinatario. Las frases de alrededor no se mueven.",
    },

    points: {
      authoring: {
        label: "Escritura",
        title: "Escribe documentos como sitios web",
        body:
          "Un documento es un árbol de componentes tipados. Si has escrito React, ya " +
          "conoces esta forma: props, composición, archivos pequeños. Así que una " +
          "persona con perfil de frontend es productiva la primera tarde en lugar de " +
          "tener que aprender antes un lenguaje de plantillas.",
      },
      changeControl: {
        label: "Control de cambios",
        title: "Los documentos viven en tu repositorio",
        body:
          "Como un documento es código fuente, cambiar una frase es un pull request: " +
          "con su diff, revisado, y todavía atribuible un año después cuando " +
          "alguien pregunte quién modificó la redacción sobre los impagos. Las " +
          "pruebas afirman que un documento genera lo que esperas, así que la CI " +
          "detecta el error antes que quien lo recibe.",
      },
    },

    pullQuote:
      "Un documento es un árbol de componentes pequeños, renderizado por algo que " +
      "entiende de papel. Tú obtienes la ergonomía de un modelo de componentes. La " +
      "persona del otro lado obtiene un archivo de Word.",

    attribution: {
      name: "Liam",
      role: "Autor, Docxcelerate",
      portraitAlt: "Liam, el autor de Docxcelerate",
    },

    engine: {
      eyebrow: "El motor",
      title: "Publica una vez. Escala cuanto haga falta.",
      bodyOne:
        "El motor es donde los documentos se escriben de verdad. Rellena tus datos, " +
        "ejecuta la IA y devuelve el documento terminado. Cualquier tipo de nodo puede " +
        "usar IA, no solo los párrafos. La respuesta del modelo o bien se convierte en " +
        "el texto, escrito a partir de la información que le das, o bien toma una " +
        "decisión de la que depende el documento.",
      bodyTwo:
        "Publicas una plantilla en el motor una sola vez. A partir de ahí, cualquier " +
        "sistema puede llamar a su API con un conjunto de datos y recibir un documento. " +
        "Hay un motor gratuito que puedes alojar tú. La nube gestionada ejecuta el " +
        "completo, con mucho que la versión gratuita no tiene, y llegará pronto.",
      ctaHow: "Cómo funciona el motor →",
      ctaCloud: "Nube gestionada",
      /** {count} is a number of documents. */
      volume: "{count} informes este trimestre",
      steps: {
        build: {
          title: "Compilar",
          detail:
            "El framework convierte tu documento en un paquete. Este paso se ejecuta en tu máquina.",
        },
        publish: {
          title: "Publicar",
          detail: "Envías el paquete a un motor. El motor lo guarda y le da un nombre.",
        },
        write: {
          title: "Escribir",
          detail:
            "Tu aplicación envía un conjunto de datos. El motor devuelve el documento terminado.",
        },
      },
    },



    docs: {
      eyebrow: "Documentación",
      title: "Todo está documentado",
      lead:
        "Cada tipo de nodo, cada flag de la CLI y cada archivo que escribe una " +
        "compilación, con vistas previas generadas por el renderizador real, de modo " +
        "que nada en la página pueda describir un helper que ya no existe.",
      ctaAll: "Documentación completa →",
      cards: {
        startHere: {
          title: "Empieza aquí",
          blurb: "Crea un workspace y abre tu primer documento.",
        },
        documentsAndNodes: {
          title: "Documentos y nodos",
          blurb: "El modelo de componentes con el que se construye un documento.",
        },
        nodeModel: {
          title: "El modelo de nodos",
          blurb: "Cada tipo de nodo, con una vista previa de cada uno.",
        },
        writingNodes: {
          title: "Escribir nodos",
          blurb: "Las piezas de las que se compone un documento.",
        },
        cli: {
          title: "Comandos de la CLI",
          blurb: "dxcl init, new, node.",
        },
        entrypoints: {
          title: "Puntos de entrada del paquete",
          blurb: "Todo lo importable y lo que exporta.",
        },
      },
    },

  },

  demo: {
    showFiles: "Mostrar los archivos del workspace",
    hideFiles: "Ocultar los archivos del workspace",
    workspaceFiles: "Archivos del workspace",
    resize: "Repartir el workspace y la vista previa",
    resizeHint: "Arrastra para repartir · doble clic para restablecer",
    zoom: "Zoom de la vista previa",
    zoomIn: "Acercar",
    zoomOut: "Alejar",
    fit: "Ajustar",
    frameTitle: "El documento seleccionado, renderizado por Docxcelerate",
    caption:
      "Una vista previa real del documento renderizándose, sin capturas de pantalla ni trucos ocultos.",
    counts: "{documents} documentos · {files} archivos",
    sectors: {
      education: "Educación",
      housing: "Vivienda",
      insurance: "Seguros",
      consulting: "Servicios profesionales",
      finance: "Finanzas",
    },
    documents: {
      "portfolio-review": "Informe trimestral de cartera",
      "offer-of-admission": "Carta de admisión",
      "repairs-appointment": "Cita de reparación",
      "policy-renewal": "Renovación de póliza",
      invoice: "Factura",
    },
  },

  docs: {
    sidebarLabel: "Documentación",
    tocLabel: "En esta página",
    onThisPage: "En esta página",
    editOnGithub: "Editar esta página en GitHub ↗",
    viewAsMarkdown: "Leer esta página en Markdown",
    /** {version} is the release the page was built against. */
    renderedAgainst: "Renderizado con {version}",
    previous: "Anterior",
    next: "Siguiente",
    /** Names the previous/next links at the foot of a page. */
    pagination: "Página anterior y siguiente",
    untranslatedNote:
      "Esta página todavía no está traducida a tu idioma, así que se muestra en inglés.",

    groups: {
      "Start Here": "Empieza aquí",
      "Essentials": "Fundamentos",
      "Nodes": "Nodos",
      "CLI": "CLI",
      "Projects": "Proyectos",
      "Generation": "Generación",
      "Reference": "Referencia",
    },
    subgroups: {
      "Node types": "Tipos de nodo",
    },
  },

  nodes: {
    helpers: "Helpers",
    noHelpers: "Ninguno, se escribe a mano",
    kind: "Clase de nodo",
    category: "Categoría",
    resolves: "Se resuelve",
    children: "Hijos",
    option: "Opción",
    type: "Tipo",
    whatItDoes: "Qué hace",
    required: "obligatorio",
    renderNote: "Qué hacen hoy los renderizadores",
    resolvesTo: "En qué se resuelve",
    resolvesToBody:
      "El nodo tal como aparece en el {model}: el JSON que recibe un renderizador. " +
      "Sin estilos, sin maquetación.",
    endpointAsked: "Qué se le pide al endpoint",
    endpointAskedBody:
      "Resuelto con los mismos datos de ejemplo. Una compilación de vista previa se " +
      "detiene en el marcador de posición; una compilación en el momento de la " +
      "petición envía estos.",
    previewTitle: "{type}: {variant}, renderizado",
    reference: "Referencia de {title}, con vistas previas →",
    categories: {
      Structure: "Estructura",
      Text: "Texto",
      Media: "Medios",
      Data: "Datos",
    },
    status: {
      "stable": "Estable",
      "no-helper": "Aún sin helper",
      "planned": "Planeado",
    },
  },

  registry: {
    /** The detail pages, per design boards 2b and 2c. */
    backToRegistry: "Registro",
    renderedBy: "Renderizado por Docxcelerate en tiempo de compilación",
    openPreview: "Abrir ↗",
    specification: "Especificación",
    specNote: "Los valores vienen de style.page y style.typography.",
    margins: "Márgenes",
    body: "Cuerpo",
    titleLabel: "Título",
    sectionHeading: "Encabezado de sección",
    paragraphLabel: "Párrafo",
    after: "después",
    compare: "Comparar",
    compareNote: "El mismo párrafo en cada tema",
    viewing: "viendo",
    requires: "Requiere",
    requiresNothing: "Nada. Se instala solo.",
    installsAlone: "Se instala solo.",
    stepData: "Datos de vista previa",
    stepDataNote: "el JSON con el que se construyó esta vista previa",
    stepPreview: "Vista previa renderizada",
    stepPreviewNote: "el componente renderizado por su cuenta",
    stepPreviewArrangementsNote: "cada disposición renderizada por su cuenta",
    /** Marks the arrangement a component uses when none is asked for. */
    arrangementDefault: "predeterminada",
    /** Labels the picker that chooses which arrangement the preview shows. */
    arrangementLabel: "Disposición",
    /** The picker's option for showing every arrangement at once. */
    arrangementAll: "Todas las disposiciones",
    stepNodes: "Nodos resueltos",
    /** {count} is a number of nodes. */
    stepNodesNote: "lo que recibe un renderizador — {count} nodos",
    stepFiles: "Los archivos",
    stepFilesNote: "lo que escribe la instalación; el archivo seleccionado es el componente",
    dataFieldsNote: "Añádelos a tu types.ts después de instalar. Es el único paso manual.",
    copyFieldPaths: "Copiar rutas de campos",
    dataFields: "Campos de datos",
    copy: "Copiar",
    copied: "Copiado",
    copyFile: "Copiar {file}",
    /** {exports} is a list of names, {count} a number of nodes. */
    exportsResolves: "exporta {exports} · se resuelve en {count} nodos",
    categoryTags: "Categoría · etiquetas",
    nothingHere: "Aquí todavía no hay nada.",
    nothingHereNote: "El registro se construye desde el paquete. En cuanto se publique una entrada, aparecerá aquí.",
    /** The one-page registry at /registry. */
    browse: {
      title: "Registro",
      description:
        "Temas y componentes de documento listos para Docxcelerate, instalados en un proyecto con un solo comando.",
      eyebrow: "Registro",
      heading: "Registro",
      lead:
        "Temas y componentes listos para tus proyectos de documento. Instala uno con un solo comando y edítalo como si fuera código tuyo.",
      allTags: "Todo",
      themesNote: "Instalar un tema escribe document-style.ts",
      nodes: "nodos",
      noPreview: "Sin vista previa",
      /** Abre las rutas de datos de una fila; une las etiquetas filtradas. */
      readsPrefix: "lee",
      /** The page's headline; `heading` is the short name the nav uses. */
      headline: "Empieza con papel que ya funciona.",
      /** {themes} and {components} are counts. */
      counts: "{themes} temas · {components} componentes",
      useTheme: "Usar tema",
      selected: "✓ Elegido",
      add: "+ Añadir",
      added: "✓ Añadido",
      familyLabel: "Qué dibuja",
      familyNote: "Agrupados según aquello en lo que se resuelve cada componente, leído del build, no una etiqueta que alguien mantiene.",
      /** {count} is how many themes and components are picked. */
      trayCount: "Tu instalación · selección: {count}",
      trayEmpty: "elige un tema o un componente",
    },
    themes: {
      title: "Temas",
      description:
        "Temas de documento listos para Docxcelerate: tipografías, colores, " +
        "espaciado y tamaño de página, instalados con un solo comando.",
      heading: "Temas",
      lead:
        "Tipografías, colores, espaciado y página, con nombre propio. Un tema " +
        "es una importación en el proyecto y viaja con el documento generado: " +
        "quien lo renderice no necesita instalar nada.",
      cta: "Ver los temas",
    },
    components: {
      title: "Componentes",
      description:
        "Componentes de Docxcelerate ya escritos: membretes, bloques de " +
        "dirección, resúmenes de pago y más, copiados como código tuyo.",
      heading: "Componentes",
      lead:
        "Nodos que alguien ya escribió y discutió. Añadir uno copia el código " +
        "a tu proyecto y pasa a ser tuyo: sin versión, sin ruta de " +
        "actualización, nada que bifurcar.",
      cta: "Ver los componentes",
    },
    install: "Instalar",
    preview: "Vista previa",
    category: "Categoría",
    tags: "Etiquetas",
    fonts: "Tipografías",
    page: "Página",
    typography: "Tipografía",
    colours: "Colores",
    spacing: "Espaciado",
    exports: "Exporta",
    field: "Campo",
    type: "Tipo",
    whatItDoes: "Para qué sirve",
    source: "Código",
    lands: "Se copia en",
    resolvesTo: "En qué se resuelve",
    previewData: "Vista previa con",
    drawnAgainst: "Diseñado sobre el tema {theme}.",
    allThemes: "Todos los temas",
    allComponents: "Todos los componentes",
    ad: {
      heading: "Empieza tu próximo documento con ventaja.",
      body:
        "El registro reúne temas y componentes de documento. `dxcl add` copia el archivo en tu proyecto. No hay dependencia ni versión que seguir.",
      browse: "Explorar el registro",
      howItWorks: "Cómo funciona la instalación",
      written: "2 archivos escritos · ninguna dependencia añadida",
      afterInstall:
        "Tras instalar, los archivos están en tu repositorio. Edítalos como cualquier otro código.",
    },
  },

  notFound: {
    title: "Página no encontrada",
    description: "Esa página no existe.",
    heading: "Esa página no existe",
    body:
      "Puede que el enlace esté desactualizado o que la página se haya movido. La " +
      "documentación es el mejor sitio para retomar el hilo.",
    ctaDocs: "Leer la documentación →",
    ctaHome: "Volver a la página de inicio",
  },
};
