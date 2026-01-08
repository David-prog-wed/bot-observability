require("dotenv").config();

const restify = require("restify");
const {
  ActivityHandler,
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  createBotFrameworkAuthenticationFromConfiguration,
  CardFactory,
} = require("botbuilder");

/* =============================
   CONFIGURACIÓN
   ============================= */
const config = {
  MicrosoftAppId: process.env.MicrosoftAppId || "",
  MicrosoftAppPassword: process.env.MicrosoftAppPassword || "",
  MicrosoftAppTenantId: process.env.MicrosoftAppTenantId || "",
};

const credFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: config.MicrosoftAppId,
  MicrosoftAppPassword: config.MicrosoftAppPassword,
  MicrosoftAppType: "MultiTenant",
  MicrosoftAppTenantId: config.MicrosoftAppTenantId,
});

const auth = createBotFrameworkAuthenticationFromConfiguration(
  null,
  credFactory
);
const adapter = new CloudAdapter(auth);

adapter.onTurnError = async (context, error) => {
  console.error("❌ Bot error:", error);
  await context.sendActivity(
    "Hubo un error procesando tu mensaje. Escribe **menu** para reiniciar el flujo."
  );
};

/* =============================
   CONTACTOS
   ============================= */
const CONTACTS = {
  softtek_l1: {
    guardia_p1:
      "🚨 **Protocolo de Guardia Softtek (P1)**: +57 300 000 0000 | Líder puente: +57 301 000 0000",
    soporte:
      "👤 **Mesa de Ayuda Softtek**: Teams @SoporteSofttek | Correo: soporte@softtek.com",
  },
  basis_sap: {
    name: "Especialista Basis Softtek (L2)",
    contact:
      "🔧 **Basis Softtek (L2)**: Teams @BasisSofttek | Correo: basis@softtek.com",
    icon: "🔧",
  },
  infra: {
    name: "Especialista Infra Softtek (L2)",
    contact:
      "🖥️ **Infra Softtek (L2)**: Teams @InfraSofttek | Correo: infra@softtek.com",
    icon: "🖥️",
  },
  l3_sap: {
    name: "Líder SAP Softtek (L3)",
    contact:
      "👔 **Líder SAP Softtek (L3)**: Juan Pérez | Tel: +57 300 333 3333",
    icon: "👔",
  },
  l3_infra: {
    name: "Líder Infra Softtek (L3)",
    contact:
      "👔 **Líder Infra Softtek (L3)**: María González | Tel: +57 300 444 4444",
    icon: "👔",
  },
};

const SYSTEMS = {
  SAP: "sap",
  INFRA: "infra",
  OTRO: "otro",
};

const SYMPTOMS = {
  FAILOVER: "failover_cluster",
  CAIDO: "caido",
  ENCOLAMIENTO: "encolamiento",
  LENTO: "lento",
  ERRORES: "errores",
};

const ENV_ALIASES = {
  produccion: "produccion",
  prod: "produccion",
  prd: "produccion",
  qa: "qa",
  test: "qa",
  testing: "qa",
  desarrollo: "dev",
  dev: "dev",
};

/* =============================
   ESTADO (por usuario y con TTL)
   ============================= */
const DRAFT_TTL_MS = 30 * 60 * 1000; // 30 min
const incidentDrafts = new Map();

function nowMs() {
  return Date.now();
}

function purgeDrafts() {
  const t = nowMs();
  for (const [k, v] of incidentDrafts.entries()) {
    if (!v || !v.updatedAt || t - v.updatedAt > DRAFT_TTL_MS) {
      incidentDrafts.delete(k);
    }
  }
}

function draftKey(context) {
  const convId = context?.activity?.conversation?.id || "no-conv";
  const fromId = context?.activity?.from?.id || "no-from";
  return `${convId}:${fromId}`;
}

function getDraft(context) {
  purgeDrafts();
  const key = draftKey(context);
  const entry = incidentDrafts.get(key);
  return entry?.data || null;
}

function setDraft(context, data) {
  purgeDrafts();
  const key = draftKey(context);
  incidentDrafts.set(key, { data, updatedAt: nowMs() });
}

function clearDraft(context) {
  const key = draftKey(context);
  incidentDrafts.delete(key);
}

/* =============================
   HELPERS
   ============================= */
function normalize(text = "") {
  return text
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

function truncate(text = "", max = 700) {
  if (!text) return "";
  return text.length > max ? `${text.slice(0, max - 3)}...` : text;
}

function fmtNowCO() {
  return new Date().toLocaleString("es-CO", { timeZone: "America/Bogota" });
}

function detectEnv(text = "") {
  const t = normalize(text);
  for (const key of Object.keys(ENV_ALIASES)) {
    const re = new RegExp(`\\b${key}\\b`, "i");
    if (re.test(t)) return ENV_ALIASES[key];
  }
  return "produccion";
}

function extractNodeInfo(text = "") {
  const nodeMatch = text.match(/VAB\d+|VEP\s*\d+|Instance\s*\d+|Nodo\s*\d+/i);
  return nodeMatch ? nodeMatch[0].replace(/\s+/g, "") : null;
}

function extractTimestamp(text = "") {
  const m1 = text.match(
    /\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s*[AP]M/i
  );
  if (m1) return m1[0];

  const m2 = text.match(
    /\d{1,2}\/\d{1,2}\/\d{4}[,\s]+\d{1,2}:\d{2}:\d{2}\s*(a\.?\s*m\.?|p\.?\s*m\.?)?/i
  );
  return m2 ? m2[0] : null;
}

function isImageAttachment(att) {
  if (!att) return false;
  const ct = (att.contentType || "").toLowerCase();
  return ct.startsWith("image/") || ct === "application/octet-stream";
}

function symptomLabel(system, symptom) {
  if (symptom === SYMPTOMS.FAILOVER) return "🚨 FAILOVER DE CLUSTER (CRÍTICO)";

  if (system === SYSTEMS.SAP) {
    switch (symptom) {
      case SYMPTOMS.CAIDO:
        return "SAP caído / no responde";
      case SYMPTOMS.ENCOLAMIENTO:
        return "Encolamientos (SMQ1/SMQ2 / qRFC)";
      case SYMPTOMS.LENTO:
        return "SAP lento / degradado";
      case SYMPTOMS.ERRORES:
        return "Errores / dumps";
      default:
        return symptom || "N/A";
    }
  }

  if (system === SYSTEMS.INFRA) {
    switch (symptom) {
      case SYMPTOMS.CAIDO:
        return "Cluster / Servicio no responde (caído)";
      case SYMPTOMS.LENTO:
        return "Degradación / Lento";
      case SYMPTOMS.ERRORES:
        return "Errores";
      default:
        return symptom || "N/A";
    }
  }

  return symptom || "N/A";
}

function classifySeverity({ system, symptom, env }) {
  const isProd = env === "produccion";
  if (symptom === SYMPTOMS.FAILOVER) return "p1";
  if (system === SYSTEMS.SAP && symptom === SYMPTOMS.ENCOLAMIENTO && isProd)
    return "p1";
  if (system === SYSTEMS.SAP && symptom === SYMPTOMS.CAIDO && isProd)
    return "p1";
  if (system === SYSTEMS.INFRA && symptom === SYMPTOMS.CAIDO && isProd)
    return "p1";
  return "p2";
}

function coerceSystemBySymptom(system, symptom) {
  if (symptom === SYMPTOMS.ENCOLAMIENTO) return SYSTEMS.SAP;
  return system;
}

/* =============================
   DETECCIÓN
   ============================= */
function detectIncident(text = "") {
  const t = normalize(text);
  const original = text;

  const hasQueue = /\b(encolamiento|cola|smq1|smq2|qrfc)\b/.test(t);
  const isFailover = /\b(failover|fail\s*over)\b/.test(t);
  const isClusterAlert = /\b(alerta\s*cluster|cluster.*alerta)\b/.test(t);
  const isInfra = /\b(vpp|vep|cluster|infra|infraestructura)\b/.test(t);
  const isSap =
    /\b(sap|smq1|smq2|sm37|st22|sm21|sm50|sm58|sm59|sm66|st03n)\b/.test(t);
  const hasVAB = /\bvab\d+\b/.test(t);

  let system = null;
  let symptom = null;
  const env = detectEnv(text);
  const node = extractNodeInfo(original);
  const timestamp = extractTimestamp(original);

  if (isFailover && (isInfra || hasVAB || isClusterAlert)) {
    system = SYSTEMS.INFRA;
    symptom = SYMPTOMS.FAILOVER;
    return { system, symptom, env, isCritical: true, node, timestamp };
  }

  if (hasQueue || isSap) system = SYSTEMS.SAP;
  else if (isInfra) system = SYSTEMS.INFRA;

  const hasDown =
    /\b(caido|down|no responde|fuera de servicio|inaccesible)\b/.test(t);
  const hasSlow = /\b(lento|degradado|latencia|degradacion)\b/.test(t);
  const hasErrors =
    /\b(error|errores|dump|st22|timeout|fallo|exception)\b/.test(t);

  if (hasQueue) symptom = SYMPTOMS.ENCOLAMIENTO;
  else if (hasDown) symptom = SYMPTOMS.CAIDO;
  else if (hasSlow) symptom = SYMPTOMS.LENTO;
  else if (hasErrors) symptom = SYMPTOMS.ERRORES;

  const coercedSystem = coerceSystemBySymptom(system, symptom);

  const isCritical =
    symptom === SYMPTOMS.FAILOVER ||
    (coercedSystem === SYSTEMS.INFRA &&
      symptom === SYMPTOMS.CAIDO &&
      env === "produccion") ||
    (coercedSystem === SYSTEMS.SAP &&
      (symptom === SYMPTOMS.CAIDO || symptom === SYMPTOMS.ENCOLAMIENTO) &&
      env === "produccion");

  return { system: coercedSystem, symptom, env, isCritical, node, timestamp };
}

/* =============================
   RUNBOOKS
   ============================= */
function buildRunbook({ system, symptom, env }) {
  const sev = classifySeverity({ system, symptom, env });
  const step = (title, bullets = []) => ({ title, bullets });

  const generic = {
    title: "Runbook L1 (Genérico)",
    quick: [
      "Confirmar **alcance** e **impacto**.",
      "Recolectar evidencia mínima (texto alerta, hora, nodo, síntoma).",
      "Escalar a L2 con contexto (y Guardia si P1).",
    ],
    steps: [
      step("Triage rápido (2-5 min)", [
        "Confirmar si el servicio responde (si aplica).",
        "Identificar si el impacto es total o parcial.",
      ]),
      step("Evidencia mínima", [
        "Texto de la alerta completo.",
        "Hora exacta (CO) del inicio.",
        "Nodo/instancia si aparece.",
      ]),
    ],
    nextAction:
      sev === "p1"
        ? "Escalar a L2 y activar Guardia P1."
        : "Escalar a L2 si persiste.",
  };

  if (system === SYSTEMS.INFRA && symptom === SYMPTOMS.FAILOVER) {
    return {
      title: "Runbook L1 - Failover de Cluster (Infra)",
      quick: [
        "Confirmar **impacto** (total/parcial) y **inicio**.",
        "Validar si hay degradación secundaria (timeouts).",
        "PROD + impacto alto → **P1**.",
      ],
      steps: [
        step("1) Confirmar impacto", [
          "¿Usuarios sin acceso total o parcial?",
          "¿Procesos críticos afectados?",
        ]),
        step("2) Verificación rápida", [
          "Validar endpoint/URL si aplica.",
          "Si no responde → Infra L2.",
        ]),
        step("3) Escalamiento", [
          "Compartir hora del failover y nodo si se conoce.",
          "Activar Guardia P1 si PROD + impacto alto.",
        ]),
      ],
      nextAction: "Contactar Infra Softtek (L2).",
    };
  }

  if (system === SYSTEMS.INFRA && symptom === SYMPTOMS.CAIDO) {
    return {
      title: "Runbook L1 - Caída de Servicio/Cluster (Infra)",
      quick: [
        "Confirmar si es caída total.",
        "Validar endpoint/ruta si aplica.",
        "PROD + no responde → P1.",
      ],
      steps: [
        step("1) Alcance", [
          "¿Todos los usuarios o solo un grupo?",
          "¿Desde cuándo?",
        ]),
        step("2) Validación rápida", [
          "Endpoint/URL responde: sí/no.",
          "Si no responde → Infra L2.",
        ]),
        step("3) Escalamiento", [
          "Enviar evidencia y severidad.",
          "Activar Guardia si P1.",
        ]),
      ],
      nextAction: "Contactar Infra Softtek (L2).",
    };
  }

  if (system === SYSTEMS.SAP && symptom === SYMPTOMS.ENCOLAMIENTO) {
    return {
      title: "Runbook L1 - Encolamientos SAP (SMQ/qRFC)",
      quick: [
        "Identificar cola(s), backlog y **primer error**.",
        "Validar destino RFC si aplica (SM59).",
        "PROD + backlog creciendo → puede ser P1.",
      ],
      steps: [
        step("1) SMQ1/SMQ2", [
          "Nombre de cola y backlog.",
          "Primer error exacto.",
        ]),
        step("2) SM59 (si aplica)", [
          "Probar destino RFC.",
          "Registrar resultado del test.",
        ]),
        step("3) Escalamiento", [
          "Enviar backlog + primer error + hora inicio.",
          "Basis L2.",
        ]),
      ],
      nextAction: "Contactar Basis Softtek (L2).",
    };
  }

  if (system === SYSTEMS.SAP && symptom === SYMPTOMS.CAIDO) {
    return {
      title: "Runbook L1 - SAP Caído / No responde",
      quick: [
        "Definir canal afectado (GUI/Web/RFC).",
        "Buscar errores (SM21/ST22).",
        "PROD + caída total → P1.",
      ],
      steps: [
        step("1) Alcance", ["GUI/Web/RFC", "¿Mandante o todo el sistema?"]),
        step("2) Evidencia", ["SM21 logs", "ST22 dumps (si disponible)."]),
        step("3) Escalamiento", [
          "Enviar evidencia a Basis L2.",
          "Activar Guardia si P1.",
        ]),
      ],
      nextAction: "Contactar Basis Softtek (L2).",
    };
  }

  return generic;
}

/* =============================
   REPORTE EJECUTIVO
   ============================= */
function generateExecutiveSummary(draft, assignedTo = null) {
  const { system, symptom, env, node, timestamp, alertText, detectedAt } =
    draft || {};
  const sev = classifySeverity({ system, symptom, env });
  const sevIcon = sev === "p1" ? "🚨" : "⚠️";
  const detectedStr = detectedAt || fmtNowCO();
  const runbook = buildRunbook({ system, symptom, env });

  let s = `${sevIcon} **ESCALAMIENTO INTERNO SOFTTEK - ${sev.toUpperCase()}**\n\n`;
  s += `**Sistema**: ${String(system || "N/A").toUpperCase()}\n\n`;
  s += `**Evento**: ${symptomLabel(system, symptom)}\n\n`;
  s += `**Entorno**: ${String(env || "N/A").toUpperCase()}\n\n`;
  if (node) s += `**Nodo/Instancia**: ${node}\n\n`;
  if (timestamp) s += `**Hora Alerta**: ${timestamp}\n\n`;
  s += `**Detección L1**: ${detectedStr}\n\n`;
  if (assignedTo) s += `**Dirigido a**: ${assignedTo}\n\n`;

  s += `**Resumen Operativo L1**:\n`;
  s += `- Severidad: **${sev.toUpperCase()}**\n`;
  s += `- Acción sugerida: ${runbook.nextAction}\n\n`;

  if (alertText) {
    s += `**Evidencia Técnica (texto recibido)**:\n`;
    s += `\`\`\`\n${truncate(alertText, 900)}\n\`\`\`\n\n`;
  }

  s += `**Runbook L1 Aplicado**: ${runbook.title}\n\n`;
  s += `**Checklist (rápido)**:\n`;
  for (const q of runbook.quick || []) s += `- ${q}\n`;

  s += `\n---\n_Generado por Bot Softtek Observabilidad_`;
  return s;
}

/* =============================
   TARJETAS
   ============================= */
function createWelcomeCard() {
  return CardFactory.adaptiveCard({
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: "🤖 Bot de Observabilidad",
        size: "Large",
        weight: "Bolder",
        color: "Accent",
      },
      {
        type: "TextBlock",
        text: "Asistente L1. Por favor, **pega el texto de la alerta** para iniciar el diagnóstico automático.",
        wrap: true,
        spacing: "Medium",
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "📋 Reportar Manualmente",
        data: { action: "reportar_incidente" },
      },
      {
        type: "Action.Submit",
        title: "🚨 Activar Guardia P1",
        data: { action: "escalar_p1" },
      },
      { type: "Action.Submit", title: "❓ Ayuda", data: { action: "ayuda" } },
    ],
  });
}

function createHelpCard() {
  return CardFactory.adaptiveCard({
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: "📖 Guía de Operación - Bot L1",
        size: "Large",
        weight: "Bolder",
        color: "Accent",
      },
      {
        type: "TextBlock",
        text: "Este bot automatiza el triage inicial basándose en el **texto** de las alertas.",
        wrap: true,
        spacing: "Small",
      },
      {
        type: "TextBlock",
        text: "🚀 ¿Cómo usarlo?",
        weight: "Bolder",
        spacing: "Medium",
        color: "Accent",
      },
      {
        type: "TextBlock",
        text: "1. **Pega la Alerta**: Copia el texto del correo o monitoreo. El bot extraerá el sistema, síntoma y severidad automáticamente.",
        wrap: true,
        spacing: "Small",
      },
      {
        type: "TextBlock",
        text: "2. **Sigue el Runbook**: El bot te dará pasos de validación rápida según el tipo de falla detectada.",
        wrap: true,
        spacing: "Small",
      },
      {
        type: "TextBlock",
        text: "3. **Escalamiento**: Una vez validado, el bot genera un reporte ejecutivo para el especialista L2.",
        wrap: true,
        spacing: "Small",
      },
      {
        type: "TextBlock",
        text: "🛠️ Capacidades Principales",
        weight: "Bolder",
        spacing: "Medium",
        color: "Accent",
      },
      {
        type: "TextBlock",
        text: "• **Triage Inteligente**: Clasifica entre SAP e Infraestructura automáticamente.",
        wrap: true,
        spacing: "Small",
      },
      {
        type: "TextBlock",
        text: "• **Reporte Ejecutivo**: Genera un resumen técnico listo para el especialista L2/L3.",
        wrap: true,
        spacing: "Small",
      },
      {
        type: "TextBlock",
        text: "• **Escalamiento Seguro**: Gestiona contactos de guardia y requiere códigos de autorización para niveles críticos.",
        wrap: true,
        spacing: "Small",
      },
      {
        type: "TextBlock",
        text: "⚠️ Nota: El bot **no procesa imágenes**. Por favor, usa siempre el **texto** de la alerta.",
        isSubtle: true,
        wrap: true,
        spacing: "Medium",
        color: "Attention",
      },
      {
        type: "TextBlock",
        text: "💡 Tip: Escribe **'menu'** en cualquier momento para reiniciar.",
        wrap: true,
        spacing: "Small",
        isSubtle: true,
        horizontalAlignment: "Center",
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "🔙 Volver al menú",
        data: { action: "menu" },
      },
    ],
  });
}

function createSystemSelectionCard(note = null) {
  return CardFactory.adaptiveCard({
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: "📋 Reportar Incidente - Paso 1/3",
        size: "Large",
        weight: "Bolder",
        color: "Accent",
      },
      ...(note
        ? [
            {
              type: "TextBlock",
              text: `ℹ️ ${note}`,
              wrap: true,
              spacing: "Small",
              isSubtle: true,
            },
          ]
        : []),
      {
        type: "TextBlock",
        text: "¿Qué sistema está afectado?",
        wrap: true,
        spacing: "Medium",
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "Infraestructura / Cluster",
        data: { action: "select_system", system: SYSTEMS.INFRA },
      },
      {
        type: "Action.Submit",
        title: "SAP",
        data: { action: "select_system", system: SYSTEMS.SAP },
      },
      {
        type: "Action.Submit",
        title: "Otro",
        data: { action: "select_system", system: SYSTEMS.OTRO },
      },
      { type: "Action.Submit", title: "🔙 Cancelar", data: { action: "menu" } },
    ],
  });
}

function createEnvironmentSelectionCard(system) {
  return CardFactory.adaptiveCard({
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: "📋 Reportar Incidente - Paso 2/3",
        size: "Large",
        weight: "Bolder",
        color: "Accent",
      },
      {
        type: "FactSet",
        facts: [
          { title: "Sistema:", value: String(system || "").toUpperCase() },
        ],
        spacing: "Small",
      },
      {
        type: "TextBlock",
        text: "¿En qué entorno ocurre el problema?",
        wrap: true,
        spacing: "Medium",
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "🔴 Producción",
        data: { action: "select_env", env: "produccion" },
      },
      {
        type: "Action.Submit",
        title: "🟡 QA / Testing",
        data: { action: "select_env", env: "qa" },
      },
      {
        type: "Action.Submit",
        title: "🟢 Desarrollo",
        data: { action: "select_env", env: "dev" },
      },
      {
        type: "Action.Submit",
        title: "🔙 Volver",
        data: { action: "reportar_incidente" },
      },
    ],
  });
}

function getSymptomsBySystem(system) {
  if (system === SYSTEMS.INFRA) {
    return [
      { label: "🚨 Failover de Cluster (crítico)", value: SYMPTOMS.FAILOVER },
      { label: "🔴 Caído / no responde", value: SYMPTOMS.CAIDO },
      { label: "🐌 Lento / degradado", value: SYMPTOMS.LENTO },
      { label: "⚠️ Errores", value: SYMPTOMS.ERRORES },
    ];
  }
  if (system === SYSTEMS.SAP) {
    return [
      { label: "🔴 Caído / no responde", value: SYMPTOMS.CAIDO },
      {
        label: "📦 Encolamientos (SMQ1/SMQ2 / qRFC)",
        value: SYMPTOMS.ENCOLAMIENTO,
      },
      { label: "🐌 Lento / degradado", value: SYMPTOMS.LENTO },
      { label: "⚠️ Errores / dumps", value: SYMPTOMS.ERRORES },
    ];
  }
  return [
    { label: "🔴 Caído", value: SYMPTOMS.CAIDO },
    { label: "🐌 Lento", value: SYMPTOMS.LENTO },
    { label: "⚠️ Errores", value: SYMPTOMS.ERRORES },
  ];
}

function createSymptomSelectionCard(system, env) {
  const symptoms = getSymptomsBySystem(system);
  return CardFactory.adaptiveCard({
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: "📋 Reportar Incidente - Paso 3/3",
        size: "Large",
        weight: "Bolder",
        color: "Accent",
      },
      {
        type: "FactSet",
        facts: [
          { title: "Sistema:", value: String(system || "").toUpperCase() },
          { title: "Entorno:", value: String(env || "").toUpperCase() },
        ],
        spacing: "Small",
      },
      {
        type: "TextBlock",
        text: "¿Cuál es el síntoma principal?",
        wrap: true,
        spacing: "Medium",
      },
    ],
    actions: symptoms
      .map((s) => ({
        type: "Action.Submit",
        title: s.label,
        data: { action: "select_symptom", symptom: s.value },
      }))
      .concat([
        {
          type: "Action.Submit",
          title: "🔙 Volver",
          data: { action: "select_system", system },
        },
      ]),
  });
}

function createIncidentSummaryCard(draft) {
  const { system, env, symptom, node, timestamp, l3Enabled } = draft;
  const sev = classifySeverity({ system, symptom, env });
  const sevIcon = sev === "p1" ? "🚨" : "⚠️";

  const facts = [
    { title: "Sistema:", value: String(system || "").toUpperCase() },
    { title: "Entorno:", value: String(env || "").toUpperCase() },
    { title: "Síntoma:", value: symptomLabel(system, symptom) },
    { title: "Severidad:", value: `${sevIcon} ${sev.toUpperCase()}` },
  ];
  if (node) facts.push({ title: "Nodo:", value: node });
  if (timestamp) facts.push({ title: "Hora evento:", value: timestamp });

  const actions = [];

  if (system === SYSTEMS.SAP) {
    actions.push({
      type: "Action.Submit",
      title: "🔧 Contactar Basis Softtek (L2)",
      data: { action: "contactar_l2_sap" },
    });
    if (l3Enabled) {
      actions.push({
        type: "Action.Submit",
        title: "👔 Contactar Líder SAP Softtek (L3)",
        data: { action: "contactar_l3_sap" },
      });
    }
  } else if (system === SYSTEMS.INFRA) {
    actions.push({
      type: "Action.Submit",
      title: "🖥️ Contactar Infra Softtek (L2)",
      data: { action: "contactar_l2_infra" },
    });
    if (l3Enabled) {
      actions.push({
        type: "Action.Submit",
        title: "👔 Contactar Líder Infra Softtek (L3)",
        data: { action: "contactar_l3_infra" },
      });
    }
  } else {
    actions.push({
      type: "Action.Submit",
      title: "📋 Completar con Reportar Incidente",
      data: { action: "reportar_incidente" },
    });
  }

  if (sev === "p1") {
    actions.push({
      type: "Action.Submit",
      title: "🚨 Activar Guardia P1",
      data: { action: "escalar_p1" },
    });
  }

  actions.push({
    type: "Action.Submit",
    title: "🔙 Menú",
    data: { action: "menu" },
  });

  const l3Note = l3Enabled
    ? "🔓 **L3 habilitado** (autorización L2 registrada)."
    : "🔒 **L3 oculto por defecto**. Se habilita solo cuando **L2 autoriza** el escalamiento.";

  return CardFactory.adaptiveCard({
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: "🧭 Análisis de Incidente Softtek",
        size: "Large",
        weight: "Bolder",
        color: "Accent",
      },
      { type: "FactSet", facts, spacing: "Medium" },
      {
        type: "TextBlock",
        text:
          sev === "p1"
            ? "⚠️ **Este incidente requiere activación de Guardia P1.**"
            : "ℹ️ Escalamiento estándar a especialista L2.",
        wrap: true,
        color: sev === "p1" ? "Attention" : "Default",
        spacing: "Medium",
      },
      {
        type: "TextBlock",
        text: "Al seleccionar un especialista, se enviará automáticamente el **Reporte Ejecutivo** con el triage realizado.",
        wrap: true,
        isSubtle: true,
        spacing: "Small",
      },
      {
        type: "TextBlock",
        text: l3Note,
        wrap: true,
        isSubtle: true,
        spacing: "Small",
      },
    ],
    actions,
  });
}

function createEscalationCard() {
  return CardFactory.adaptiveCard({
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: "🚨 Protocolo de Guardia P1",
        size: "Large",
        weight: "Bolder",
        color: "Attention",
      },
      {
        type: "TextBlock",
        text: CONTACTS.softtek_l1.guardia_p1,
        wrap: true,
        spacing: "Medium",
      },
      {
        type: "TextBlock",
        text: "Recuerda compartir el Reporte Ejecutivo generado con el equipo de guardia.",
        wrap: true,
        spacing: "Medium",
        isSubtle: true,
      },
    ],
    actions: [
      { type: "Action.Submit", title: "🔙 Menú", data: { action: "menu" } },
    ],
  });
}

function createL2AuthorizationCard(system) {
  const targetLabel =
    system === SYSTEMS.SAP
      ? "Líder SAP (L3)"
      : system === SYSTEMS.INFRA
      ? "Líder Infra (L3)"
      : "L3";

  return CardFactory.adaptiveCard({
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: "🔐 Autorización L2 para escalar a L3",
        size: "Large",
        weight: "Bolder",
        color: "Accent",
      },
      {
        type: "TextBlock",
        text:
          "Usa esto **solo si L2 lo autorizó** (casos extremos). Al confirmar, se habilita el botón para contactar a " +
          `**${targetLabel}** en la tarjeta de análisis.`,
        wrap: true,
        spacing: "Medium",
      },
      {
        type: "TextBlock",
        text: "Para confirmar, ingresa el código de autorización L2:",
        wrap: true,
        spacing: "Medium",
        weight: "Bolder",
      },
      {
        type: "Input.Text",
        id: "l2_code",
        placeholder: "Código L2",
        maxLength: 20,
        isRequired: true,
      },
      {
        type: "TextBlock",
        text: "Si estás en **Bot Framework Emulator** y no te aparece el campo, puedes escribir el código directamente en el chat.",
        wrap: true,
        spacing: "Small",
        isSubtle: true,
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: `✅ Confirmar y habilitar ${targetLabel}`,
        data: { action: "habilitar_l3" },
      },
      {
        type: "Action.Submit",
        title: "❌ Mantener solo L2",
        data: { action: "deshabilitar_l3" },
      },
      { type: "Action.Submit", title: "🔙 Menú", data: { action: "menu" } },
    ],
  });
}

function createL3SystemPickCard() {
  return CardFactory.adaptiveCard({
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: "🧩 Selección requerida para escalar a L3",
        size: "Large",
        weight: "Bolder",
        color: "Accent",
      },
      {
        type: "TextBlock",
        text: "No tengo claro el sistema del incidente (o se perdió el contexto). Selecciona a qué área corresponde el escalamiento L3:",
        wrap: true,
        spacing: "Medium",
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "👔 L3 Infra",
        data: { action: "force_l3_system", system: SYSTEMS.INFRA },
      },
      {
        type: "Action.Submit",
        title: "👔 L3 SAP (Basis)",
        data: { action: "force_l3_system", system: SYSTEMS.SAP },
      },
      { type: "Action.Submit", title: "🔙 Menú", data: { action: "menu" } },
    ],
  });
}

function createIncidentClosureCard() {
  return CardFactory.adaptiveCard({
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: "✅ Incidente atendido",
        size: "Large",
        weight: "Bolder",
        color: "Good",
      },
      {
        type: "TextBlock",
        text: "El resumen fue generado y compartido con el especialista.\n\nSi el especialista L2 te indica que debes escalar al Jefe/Líder (L3), usa el botón de abajo.",
        wrap: true,
        spacing: "Medium",
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "👔 Escalar a L3 (Requiere Código)",
        data: { action: "solicitar_l3_manual" },
      },
      {
        type: "Action.Submit",
        title: "🏠 Volver al menú",
        data: { action: "menu" },
      },
      {
        type: "Action.Submit",
        title: "➕ Reportar otro incidente",
        data: { action: "reportar_incidente" },
      },
    ],
  });
}

/* =============================
   UTIL: Reporte por destinatario
   ============================= */
function markReportSent(draft, key) {
  if (!draft) return draft;
  if (!draft.reportSentTo) draft.reportSentTo = {};
  draft.reportSentTo[key] = true;
  return draft;
}

function wasReportSent(draft, key) {
  return Boolean(draft?.reportSentTo?.[key]);
}

async function sendReportAndContact(
  context,
  assignedToName,
  contactText,
  recipientKey
) {
  const draft = getDraft(context);

  if (!draft || !draft.system || !draft.symptom || !draft.env) {
    await context.sendActivity(
      "No tengo un incidente activo para generar reporte. Pega la alerta (texto) o usa **Reportar Incidente**."
    );
    await context.sendActivity({ attachments: [createWelcomeCard()] });
    return;
  }

  if (!wasReportSent(draft, recipientKey)) {
    const report = generateExecutiveSummary(draft, assignedToName);
    await context.sendActivity(
      `📄 **Resumen Ejecutivo para ${assignedToName}:**\n\n${report}`
    );
    markReportSent(draft, recipientKey);
    setDraft(context, draft);
  }

  const sev = classifySeverity(draft);
  if (sev === "p1") {
    await context.sendActivity(
      "🚨 Nota: Clasificado como **P1**. Si el impacto es alto, **activa Guardia** y notifica según protocolo."
    );
  }

  await context.sendActivity(contactText);
  await context.sendActivity({ attachments: [createIncidentClosureCard()] });
}

/* =============================
   VALIDACIÓN DE CÓDIGO L2
   ============================= */
const VALID_L2_CODES = ["L2SOFT", "ESCALATE", "ADMIN123"];

function isValidL2Code(code) {
  if (!code) return false;
  return VALID_L2_CODES.includes(String(code).toUpperCase().trim());
}

/* =============================
   BOT PRINCIPAL
   ============================= */
class TeamsObservabilidadBot extends ActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      const rawText = (context.activity.text || "").trim();
      const normText = normalize(rawText);
      const value = context.activity.value;
      const attachments = context.activity.attachments || [];

      // Bloqueo de imágenes
      const hasImage = attachments.some(isImageAttachment);
      if (hasImage && !rawText) {
        await context.sendActivity(
          "📷 **No puedo procesar imágenes.**\n\nPara ayudarte con el triage, por favor **copia y pega el texto** de la alerta o usa el botón **Reportar Manualmente**."
        );
        await context.sendActivity({ attachments: [createWelcomeCard()] });
        return await next();
      }

      // Comandos de texto corto / Saludos
      if (
        /\b(hola|holi|buenas|buenos dias|buenas tardes|buenas noches|hey|menu|inicio)\b/.test(
          normText
        )
      ) {
        clearDraft(context);
        await context.sendActivity({ attachments: [createWelcomeCard()] });
        return await next();
      }

      if (
        /\b(gracias|muchas gracias|ok|vale|listo|perfecto|excelente|genial|👍)\b/.test(
          normText
        )
      ) {
        await context.sendActivity(
          "¡Con gusto! ✅\n\nSi necesitas algo más:\n* Pega una alerta\n* Escribe **menu**"
        );
        return await next();
      }

      const draftForTextFallback = getDraft(context);
      if (
        rawText &&
        isValidL2Code(rawText) &&
        draftForTextFallback?.awaitingL3Auth === true
      ) {
        const nextDraft = {
          ...(draftForTextFallback || {}),
          l3Enabled: true,
          awaitingL3Auth: false,
        };
        setDraft(context, nextDraft);
        await context.sendActivity(
          "✅ Autorización L2 registrada (por texto). L3 habilitado."
        );
        await context.sendActivity({
          attachments: [createIncidentSummaryCard(nextDraft)],
        });
        return await next();
      }

      if (
        normText === "l3" ||
        normText === "escalar l3" ||
        normText === "escalamiento l3"
      ) {
        const draft = getDraft(context);
        if (draft && draft.system) {
          setDraft(context, { ...(draft || {}), awaitingL3Auth: true });
          await context.sendActivity({
            attachments: [createL2AuthorizationCard(draft.system)],
          });
        } else {
          await context.sendActivity(
            "No hay un incidente activo para escalar. Escribe **menu** para iniciar uno."
          );
        }
        return await next();
      }

      if (value && value.action) {
        const draft = getDraft(context) || {};

        switch (value.action) {
          case "menu":
            clearDraft(context);
            await context.sendActivity({ attachments: [createWelcomeCard()] });
            break;

          case "ayuda":
            await context.sendActivity({ attachments: [createHelpCard()] });
            break;

          case "reportar_incidente":
            setDraft(context, {
              detectedAt: fmtNowCO(),
              l3Enabled: false,
              awaitingL3Auth: false,
            });
            await context.sendActivity({
              attachments: [createSystemSelectionCard()],
            });
            break;

          case "select_system": {
            const nextDraft = {
              ...(draft || {}),
              system: value.system,
              detectedAt: draft?.detectedAt || fmtNowCO(),
              l3Enabled: draft?.l3Enabled ?? false,
              awaitingL3Auth: draft?.awaitingL3Auth ?? false,
            };
            setDraft(context, nextDraft);
            await context.sendActivity({
              attachments: [createEnvironmentSelectionCard(nextDraft.system)],
            });
            break;
          }

          case "select_env": {
            const nextDraft = {
              ...(draft || {}),
              env: value.env,
              detectedAt: draft?.detectedAt || fmtNowCO(),
              l3Enabled: draft?.l3Enabled ?? false,
              awaitingL3Auth: draft?.awaitingL3Auth ?? false,
            };
            setDraft(context, nextDraft);
            await context.sendActivity({
              attachments: [
                createSymptomSelectionCard(nextDraft.system, nextDraft.env),
              ],
            });
            break;
          }

          case "select_symptom": {
            let sys = draft?.system;
            const sym = value.symptom;
            sys = coerceSystemBySymptom(sys, sym);
            const nextDraft = {
              ...(draft || {}),
              system: sys,
              symptom: sym,
              detectedAt: draft?.detectedAt || fmtNowCO(),
              l3Enabled: draft?.l3Enabled ?? false,
              awaitingL3Auth: draft?.awaitingL3Auth ?? false,
            };
            setDraft(context, nextDraft);
            await context.sendActivity({
              attachments: [createIncidentSummaryCard(nextDraft)],
            });
            break;
          }

          case "escalar_p1":
            await context.sendActivity({
              attachments: [createEscalationCard()],
            });
            break;

          case "solicitar_l3_manual": {
            const currentDraft = getDraft(context);
            if (currentDraft) {
              setDraft(context, {
                ...(currentDraft || {}),
                awaitingL3Auth: true,
              });
            }
            if (currentDraft && currentDraft.system) {
              await context.sendActivity({
                attachments: [createL2AuthorizationCard(currentDraft.system)],
              });
            } else if (currentDraft) {
              await context.sendActivity({
                attachments: [createL3SystemPickCard()],
              });
            } else {
              await context.sendActivity(
                "No hay un incidente activo. Escribe **menu** para iniciar uno."
              );
            }
            break;
          }

          case "force_l3_system": {
            const nextDraft = {
              ...(draft || {}),
              system: value.system,
              detectedAt: draft?.detectedAt || fmtNowCO(),
              awaitingL3Auth: true,
            };
            setDraft(context, nextDraft);
            await context.sendActivity({
              attachments: [createL2AuthorizationCard(value.system)],
            });
            break;
          }

          case "habilitar_l3": {
            const code = value?.l2_code;
            if (!isValidL2Code(code)) {
              await context.sendActivity(
                "❌ Código inválido. Pide el código al especialista L2."
              );
              await context.sendActivity({
                attachments: [createL2AuthorizationCard(draft?.system)],
              });
              break;
            }
            const nextDraft = {
              ...(draft || {}),
              l3Enabled: true,
              awaitingL3Auth: false,
            };
            setDraft(context, nextDraft);
            await context.sendActivity(
              "✅ Autorización L2 registrada. L3 habilitado."
            );
            await context.sendActivity({
              attachments: [createIncidentSummaryCard(nextDraft)],
            });
            break;
          }

          case "deshabilitar_l3": {
            const nextDraft = {
              ...(draft || {}),
              l3Enabled: false,
              awaitingL3Auth: false,
            };
            setDraft(context, nextDraft);
            await context.sendActivity("✅ OK. Solo L2.");
            await context.sendActivity({
              attachments: [createIncidentSummaryCard(nextDraft)],
            });
            break;
          }

          case "contactar_l2_sap":
            await sendReportAndContact(
              context,
              CONTACTS.basis_sap.name,
              CONTACTS.basis_sap.contact,
              "l2_sap"
            );
            break;

          case "contactar_l2_infra":
            await sendReportAndContact(
              context,
              CONTACTS.infra.name,
              CONTACTS.infra.contact,
              "l2_infra"
            );
            break;

          case "contactar_l3_sap": {
            const cur = getDraft(context);
            if (!cur?.l3Enabled) {
              await context.sendActivity(
                "🔒 L3 no habilitado. Solicita autorización primero."
              );
              await context.sendActivity({
                attachments: [createL2AuthorizationCard(cur?.system)],
              });
              break;
            }
            await sendReportAndContact(
              context,
              CONTACTS.l3_sap.name,
              CONTACTS.l3_sap.contact,
              "l3_sap"
            );
            break;
          }

          case "contactar_l3_infra": {
            const cur = getDraft(context);
            if (!cur?.l3Enabled) {
              await context.sendActivity(
                "🔒 L3 no habilitado. Solicita autorización primero."
              );
              await context.sendActivity({
                attachments: [createL2AuthorizationCard(cur?.system)],
              });
              break;
            }
            await sendReportAndContact(
              context,
              CONTACTS.l3_infra.name,
              CONTACTS.l3_infra.contact,
              "l3_infra"
            );
            break;
          }

          default:
            await context.sendActivity(
              "Acción no reconocida. Escribe **menu**."
            );
        }

        return await next();
      }

      if (rawText && rawText.length >= 5) {
        const detection = detectIncident(rawText);
        let { system, symptom, env, isCritical, node, timestamp } = detection;

        if (!system && symptom) {
          setDraft(context, {
            detectedAt: fmtNowCO(),
            env: env || "produccion",
            symptom,
            node: node || null,
            timestamp: timestamp || null,
            alertText: rawText,
            l3Enabled: false,
            awaitingL3Auth: false,
          });
          await context.sendActivity(
            "Detecté un incidente, pero no el sistema."
          );
          await context.sendActivity({
            attachments: [
              createSystemSelectionCard("No pude inferir sistema."),
            ],
          });
          return await next();
        }

        if (system && symptom) {
          system = coerceSystemBySymptom(system, symptom);
          const inferred = {
            system,
            env: env || "produccion",
            symptom,
            node: node || null,
            timestamp: timestamp || null,
            detectedAt: fmtNowCO(),
            alertText: rawText,
            l3Enabled: false,
            awaitingL3Auth: false,
          };
          setDraft(context, inferred);
          if (isCritical)
            await context.sendActivity("🚨 **ALERTA CRÍTICA DETECTADA**");
          await context.sendActivity({
            attachments: [createIncidentSummaryCard(inferred)],
          });
          return await next();
        }
      }

      await context.sendActivity(
        "No detecté un incidente claro. Escribe **menu**."
      );
      await context.sendActivity({ attachments: [createWelcomeCard()] });
      return await next();
    });

    this.onMembersAdded(async (context, next) => {
      for (const m of context.activity.membersAdded || []) {
        if (m.id !== context.activity.recipient.id) {
          await context.sendActivity({ attachments: [createWelcomeCard()] });
        }
      }
      await next();
    });
  }
}

const bot = new TeamsObservabilidadBot();

const server = restify.createServer();
server.use(restify.plugins.bodyParser());
const port = process.env.PORT || 3978;
server.listen(port, () =>
  console.log(`✅ Bot escuchando en http://localhost:${port}`)
);
server.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, (context) => bot.run(context));
});
