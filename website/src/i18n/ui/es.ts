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
    cloud: "Cloud",
    github: "GitHub",
    switchTheme: "Cambiar de tema",
    language: "Idioma",
  },

  footer: {
    label: "Pie de página",
    licence:
      "Con licencia MIT. La escritura, la vista previa y el empaquetado de DOCX se ejecutan por completo en tu máquina.",
    docs: "Docs",
    cloud: "Cloud",
    github: "GitHub",
    npm: "npm",
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
    copyInstall: "Copiar el comando de instalación",
    copy: "Copiar",
    copied: "Copiado",
    copyFallback: "Pulsa ⌘C",
    integrityTitle: "{hash} · haz clic para copiar",
    /** Accessible name for the registry marks beside the version. */
    onNpm: "Docxcelerate en npm",
    onJsr: "Docxcelerate en JSR",

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
      ai: {
        label: "IA",
        title: "IA a nivel de componente",
        body:
          "La IA entra por hooks, dentro de los componentes que ya escribes. Un " +
          "componente le da al modelo su contexto y lo que quieres que escriba, así " +
          "que produce esa única parte del documento mientras todo lo que la rodea " +
          "sigue siendo determinista. Tú decides cuánto se genera, un componente " +
          "cada vez.",
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
      "Docxcelerate trata un documento como un framework de interfaz trata una " +
      "pantalla: componentes pequeños, compuestos en un árbol, renderizados por algo " +
      "que entiende de papel. Tú obtienes la ergonomía de un modelo de componentes, y " +
      "quien lo recibe obtiene un documento de Word.",

    engine: {
      title: "El motor",
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

    scale: {
      eyebrow: "Empresa",
      title: "Pensado para los informes que ya escribes a mano",
      lead:
        "Una firma que produce el mismo informe cientos de veces al mes ya tiene la " +
        "maquetación y casi toda la redacción. Lo que cambia es la persona para la " +
        "que se escribe. Reconstruye ese informe con componentes, mantén cada " +
        "ejecución idéntica y deja al modelo solo las partes que dependen de quién " +
        "lo va a leer.",
      traits: {
        volume: {
          title: "Una plantilla, cada destinatario",
          body:
            "Publica la plantilla una vez y luego llámala por persona. Un informe o " +
            "cien mil son la misma llamada, repetida.",
        },
        determinism: {
          title: "El mismo documento siempre",
          body:
            "Todo lo que no marcaste como generado se renderiza igual en cada " +
            "ejecución. Solo pueden variar las partes que elegiste.",
        },
        integration: {
          title: "Lo llaman los sistemas que ya tienes",
          body:
            "Tu CRM, tu gestor de expedientes o tu sistema de facturación envía sus " +
            "datos y recibe un .docx. Nadie exporta una hoja de cálculo ni abre Word.",
        },
        reproducibility: {
          title: "Reproducible un año después",
          body:
            "Las plantillas están versionadas, así que cualquier documento se puede " +
            "reconstruir con la misma plantilla y los mismos datos. Una auditoría " +
            "recibe una compilación, no un archivo.",
        },
      },
    },

    openSource: {
      eyebrow: "Código abierto",
      title: "Código abierto, y hecho para seguir siéndolo",
      bodyOne:
        "El framework, los renderizadores, el modelo de nodos y la CLI tienen " +
        "licencia MIT y se desarrollan a la vista de todos. Lee el código que escribe " +
        "tus documentos, bifúrcalo o incorpóralo a tu propia compilación. La " +
        "escritura, la vista previa y el empaquetado de DOCX se ejecutan íntegramente " +
        "en tu máquina: sin cuenta, sin subidas, sin llamadas de red.",
      bodyTwo:
        "El motor de generación se puede autoalojar gratis, así que generar " +
        "documentos a escala nunca depende de que un proveedor siga en el mercado ni " +
        "de que una lista de precios no cambie. La nube gestionada es una comodidad " +
        "para equipos que prefieren no operarlo ellos mismos, no la única puerta de " +
        "entrada.",
      facts: {
        licence: {
          title: "Licencia MIT",
          body:
            "Úsalo comercialmente, modifícalo, distribúyelo. Sin recuento de puestos " +
            "ni tarifa por documento.",
        },
        local: {
          title: "Se ejecuta en tu máquina",
          body:
            "Los documentos y los datos se quedan en tu portátil y en tu CI salvo que " +
            "los envíes a un motor que hayas elegido.",
        },
        selfHost: {
          title: "Motor autoalojable",
          body: "El motor gratuito corre en tu propia infraestructura, dentro de tu propia red.",
        },
      },
      ctaSource: "Ver el código →",
      ctaNpm: "Verlo en npm",
    },

    docs: {
      eyebrow: "Documentación",
      title: "Léelo entero",
      lead:
        "Cada tipo de nodo, cada flag de la CLI y cada archivo que escribe una " +
        "compilación, con vistas previas generadas por el renderizador real, de modo " +
        "que nada en la página pueda describir un helper que ya no existe.",
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

    agentSkill: {
      eyebrow: "Agent skills",
      title: "Dale el conjunto a tu agente",
      lead:
        "Un solo archivo Markdown le enseña a un agente de código cómo se arman aquí " +
        "los documentos: el modelo de componentes, las reglas con las que tropiezan " +
        "los agentes y todos los comandos. Colócalo y pide un documento en lugar de " +
        "escribir el primero tú.",
      where: "Dónde va el archivo",
      copySkill: "Copiar la skill",
      references:
        "Es Markdown corriente, y a su lado hay cuatro archivos de referencia en " +
        "skills/docxcelerate/.",
      agents: {
        "claude-code":
          "Copia la carpeta en .claude/skills/ para un proyecto, o en " +
          "~/.claude/skills/ para todos. Se carga sola en cuanto aparece un proyecto " +
          "de documentos.",
        "cursor":
          "Guárdalo como archivo de reglas y Cursor lo lee en ese proyecto. O " +
          "menciónalo con @ en el chat cuando lo necesites.",
        "copilot":
          "Pégalo en .github/copilot-instructions.md y Copilot lo aplica en todo el " +
          "repositorio.",
        "agents-md":
          "Codex, Gemini CLI, Aider y Cline leen AGENTS.md en la raíz del repositorio. " +
          "Pégalo ahí, o enlázalo si tu agente abre archivos.",
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
    },
    documents: {
      "offer-of-admission": "Carta de admisión",
      "repairs-appointment": "Cita de reparación",
      "policy-renewal": "Renovación de póliza",
    },
  },

  docs: {
    sidebarLabel: "Documentación",
    tocLabel: "En esta página",
    onThisPage: "En esta página",
    editOnGithub: "Editar esta página en GitHub ↗",
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
