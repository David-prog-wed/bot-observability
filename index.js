// index.js
require("dotenv").config();

// Configuración de Application Insights (opcional)
/*const appInsights = require("applicationinsights");
const aiConn =
  process.env.APPLICATIONINSIGHTS_CONNECTION_STRING ||
  process.env.APPINSIGHTS_CONNECTION_STRING;

if (aiConn) {
  appInsights
    .setup(aiConn)
    .setAutoCollectConsole(true, true)
    .setSendLiveMetrics(true)
    .start();
}*/

const restify = require("restify");
const {
  ActivityHandler,
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  createBotFrameworkAuthenticationFromConfiguration,
} = require("botbuilder");

/* =============================
   Configuración y credenciales
   ============================= */
const config = {
  MicrosoftAppId: process.env.MicrosoftAppId || "",
  MicrosoftAppPassword: process.env.MicrosoftAppPassword || "",
  MicrosoftAppTenantId: process.env.MicrosoftAppTenantId || "",
};

const credFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: config.MicrosoftAppId,
  MicrosoftAppPassword: config.MicrosoftAppPassword,
  MicrosoftAppType: "SingleTenant",
  MicrosoftAppTenantId: config.MicrosoftAppTenantId,
});

const auth = createBotFrameworkAuthenticationFromConfiguration(
  null,
  credFactory
);
const adapter = new CloudAdapter(auth);

// Manejo de errores
adapter.onTurnError = async (context, error) => {
  console.error("❌ Bot error:", error);
  await context.sendActivity("Hubo un error procesando tu mensaje.");
};

/* =============================
   Datos simulados
   ============================= */
const RUNBOOKS = [
  {
    sistema: "sap",
    link: "https://tu-sharepoint/sites/Observabilidad/Runbooks/ReinicioSAP.pdf",
  },
  {
    sistema: "java",
    link: "https://tu-sharepoint/sites/Observabilidad/Runbooks/ServicioJava.pdf",
  },
  {
    sistema: "pagos",
    link: "https://tu-sharepoint/sites/Observabilidad/Runbooks/RunbookPagos.pdf",
  },
];

const DASHBOARDS = [
  { servicio: "sap", link: "https://grafana/sap" },
  { servicio: "pagos", link: "https://grafana/pagos" },
  { servicio: "auth", link: "https://grafana/auth" },
  { servicio: "db", link: "https://grafana/db" },
];

const ONCALL = {
  p1: "📞 On-Call P1: +57 300 000 0000 | Líder puente: +57 301 000 0000",
  p2: "📞 On-Call P2: +57 311 000 0000",
};

/* =============================
   Helpers
   ============================= */
function normalize(text = "") {
  return text
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

function buscarRunbooks(query = "") {
  const term = normalize(query);
  return RUNBOOKS.filter((r) => normalize(r.sistema).includes(term));
}

function dashboardLink(servicio = "") {
  const term = normalize(servicio);
  return DASHBOARDS.find((d) => normalize(d.servicio).includes(term));
}

const SYSTEM_ALIASES = {
  sap: "sap",
  pagos: "pagos",
  payment: "pagos",
  auth: "auth",
  autenticacion: "auth",
  login: "auth",
  bd: "db",
  database: "db",
  db: "db",
  grafana: "grafana",
  java: "java",
  kafka: "queue",
  rabbit: "queue",
  cola: "queue",
  colas: "queue",
  encolamiento: "queue",
  queue: "queue",
  vpp: "vpp",
  valorem: "vpp",
};

function detectSystem(text = "") {
  const norm = normalize(text);
  for (const key of Object.keys(SYSTEM_ALIASES)) {
    if (norm.includes(key)) return SYSTEM_ALIASES[key];
  }
  return null;
}

function classifySeverity(text = "") {
  const t = normalize(text);
  const isProd = /\b(produccion|prod)\b/.test(t);
  const strong =
    /\b(caido|down|no responde|fuera de servicio|critico|critica|p1)\b/.test(t);
  const moderate = /\b(lento|degradado|intermitente|p2)\b/.test(t);

  if (strong || (isProd && /(error|fallo|cola|encolamiento)/.test(t)))
    return "p1";
  if (moderate) return "p2";
  return "info";
}

function suggestChecks(text = "", system = "") {
  const t = normalize(text);
  const out = [];

  if (/encolamiento|cola|queue|kafka|rabbit/.test(t))
    out.push("Revisar colas (Kafka/Rabbit): profundidad y consumidores.");
  if (/sap/.test(t) || system === "sap")
    out.push("Validar conexión SAP (RFC/ping) y jobs en SM37.");
  if (/auth|login/.test(t) || system === "auth")
    out.push("Revisar tiempos de respuesta y errores 401/500.");
  if (/pagos|payment/.test(t) || system === "pagos")
    out.push("Correlacionar errores por medio de pago.");
  if (/java/.test(t) || system === "java")
    out.push("Ver heap/GC y errores en logs (OutOfMemory, timeouts).");
  if (/bd|database|db/.test(t) || system === "db")
    out.push("Chequear locks y conexiones activas en BD.");
  if (/grafana/.test(t) || system === "grafana")
    out.push("Ver estado de datasource y paneles clave.");
  if (/vpp/.test(t) || system === "vpp")
    out.push("Validar microservicios VPP y colas asociadas.");

  out.push("Revisar panel en Grafana y últimos despliegues.");
  return out;
}

function buildIncidentResponse({ text }) {
  const system = detectSystem(text) || "desconocido";
  const sev = classifySeverity(text);
  const checks = suggestChecks(text, system);

  const rb = buscarRunbooks(system)[0];
  const dash = dashboardLink(system);

  const lines = [];
  lines.push(`🧭 **Análisis inicial**`);
  lines.push(`• Sistema: **${system}**`);
  lines.push(`• Severidad sugerida: **${sev.toUpperCase()}**\n`);
  if (dash) lines.push(`📊 Dashboard: ${dash.link}`);
  if (rb) lines.push(`📖 Runbook: ${rb.link}\n`);

  lines.push("✅ **Sugerencias rápidas:**");
  for (const s of checks.slice(0, 5)) lines.push(`- ${s}`);

  if (sev === "p1") {
    lines.push("\n🚨 Se recomienda escalar como **P1**.");
    lines.push(ONCALL.p1);
  } else if (sev === "p2") {
    lines.push("\n⚠️ Puede tratarse de **P2**, monitorear impacto.");
    lines.push(ONCALL.p2);
  } else {
    lines.push("\nℹ️ Si empeora el impacto, considera escalar.");
  }

  return lines.join("\n");
}

/* =============================
   Bot principal
   ============================= */
class TeamsObservabilidadBot extends ActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      const rawText = (context.activity.text || "").trim();
      const normText = normalize(rawText);

      console.log(`💬 Usuario: "${rawText}" -> "${normText}"`);

      // 1) Saludos
      if (
        /\b(hola|holi|buenas|buenos dias|buenas tardes|buenas noches|hey)\b/.test(
          normText
        )
      ) {
        await context.sendActivity(
          "¡Hola! 👋 Soy el bot de Observabilidad. Puedo ayudarte con runbooks, escalamiento y dashboards. Escribe **ayuda** para ver ejemplos."
        );

        // 2) Ayuda / menú
      } else if (
        ["ayuda", "help", "?", "comandos", "menu"].includes(normText)
      ) {
        await context.sendActivity(
          [
            "🤖 **¿Cómo puedo ayudarte?**",
            "",
            "**Comandos rápidos:**",
            "- `runbook <sistema>` → Ej: `runbook sap`, `runbook pagos`",
            "- `dashboard <servicio>` → Ej: `dashboard pagos`, `dashboard auth`",
            "- `escalar p1` / `escalar p2` → Números on-call",
            "",
            "**Incidentes en lenguaje natural (yo los interpreto):**",
            "- `tengo problema con encolamiento en vpp en producción`",
            "- `usuarios no pueden loguearse; auth con errores 500`",
            "- `sap muy lento luego del despliegue`",
          ].join("\n")
        );

        // 3) Runbooks
      } else if (normText.startsWith("runbook")) {
        const query = normText.replace("runbook", "").trim();
        const results = buscarRunbooks(query);
        if (results.length > 0) {
          for (const r of results)
            await context.sendActivity(
              `📘 Runbook de **${r.sistema}** → ${r.link}`
            );
        } else {
          await context.sendActivity(
            `❌ No encontré runbook para "${query || "ese sistema"}".`
          );
        }

        // 4) Escalamientos
      } else if (normText.includes("escalar p1")) {
        await context.sendActivity(`🚨 Escalamiento P1:\n${ONCALL.p1}`);
      } else if (normText.includes("escalar p2")) {
        await context.sendActivity(`⚠️ Escalamiento P2:\n${ONCALL.p2}`);

        // 5) Dashboards
      } else if (normText.startsWith("dashboard")) {
        const servicio = normText.replace("dashboard", "").trim();
        const dash = dashboardLink(servicio);
        if (dash)
          await context.sendActivity(
            `📊 Dashboard de **${servicio}** → ${dash.link}`
          );
        else
          await context.sendActivity(
            `❌ No encontré dashboard para "${servicio || "ese servicio"}".`
          );

        // 6) Agradecimientos / despedida (opcional)
      } else if (
        /\b(gracias|muchas gracias|bye|adios|hasta luego)\b/.test(normText)
      ) {
        await context.sendActivity(
          "¡Con gusto! Si necesitas algo más, aquí estaré. 🙌"
        );

        // 7) Texto libre (incidentes) – usa el analizador
      } else if (rawText.length >= 6) {
        const suggestion = buildIncidentResponse({ text: rawText });
        await context.sendActivity(suggestion);

        // 8) Fallback
      } else {
        await context.sendActivity(
          "No te entendí del todo 🤔. Escribe **ayuda** para ver ejemplos, o describe el incidente (sistema afectado, síntoma, entorno) y te guío. 🙂"
        );
      }

      await next();
    });

    this.onMembersAdded(async (context, next) => {
      await context.sendActivity(
        "¡Hola! Estoy listo para ayudarte con runbooks, escalamiento y dashboards."
      );
      await next();
    });
  }
}

const bot = new TeamsObservabilidadBot();

/* =============================
   Servidor REST
   ============================= */
const server = restify.createServer();
server.use(restify.plugins.bodyParser());
const port = process.env.PORT || 3978;
server.listen(port, () =>
  console.log(`✅ Bot escuchando en http://localhost:${port}`)
);
server.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, (context) => bot.run(context));
});
