import { definePluginEntry } from "openclaw/plugin-sdk/plugin-entry";

const PLUGIN_ID = "microsoft365";
const DEFAULT_CONNECTION_NAME = "openclaw-microsoft365";

const configSchema = {
  type: "object",
  additionalProperties: false,
  properties: {
    account: {
      type: "string",
      description: "Default Microsoft 365 account UPN. Example: user@example.com."
    },
    tenantId: {
      type: "string",
      description: "Microsoft Entra tenant id or verified domain. Can also be provided through M365_TENANT_ID."
    },
    clientId: {
      type: "string",
      description: "Microsoft Entra application/client id. Can also be provided through M365_CLIENT_ID."
    },
    connectionName: {
      type: "string",
      description: "Optional CLI for Microsoft 365 connection name. Defaults to openclaw-microsoft365."
    }
  }
};

function cleanString(value) {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function resolveEnvReference(value) {
  if (typeof value !== "string") return value;
  const raw = value.trim();
  if (!raw) return undefined;

  const envRef = raw.match(/^\$\{([A-Z][A-Z0-9_]*)\}$/) || raw.match(/^\$([A-Z][A-Z0-9_]*)$/);
  if (envRef) return process.env[envRef[1]];

  const legacyEnvRef = raw.match(/^(?:secretref-env:|__env__:)([A-Z][A-Z0-9_]*)$/);
  if (legacyEnvRef) return process.env[legacyEnvRef[1]];

  return raw;
}

function configString(config, key, envKey, defaultValue) {
  return (
    cleanString(resolveEnvReference(config?.[key])) ??
    cleanString(process.env[envKey]) ??
    defaultValue
  );
}

export default definePluginEntry({
  id: PLUGIN_ID,
  name: "Microsoft 365",
  description: "Adds Microsoft 365 workflow guidance using the official CLI for Microsoft 365 as backend.",
  configSchema,
  register(api) {
    api.on("resolve_exec_env", async () => {
      const cfg = api.pluginConfig ?? {};
      const env = {};

      const tenantId = configString(cfg, "tenantId", "M365_TENANT_ID");
      const clientId = configString(cfg, "clientId", "M365_CLIENT_ID");
      const account = configString(cfg, "account", "M365_ACCOUNT");
      const connectionName = configString(cfg, "connectionName", "M365_CONNECTION_NAME", DEFAULT_CONNECTION_NAME);

      // CLI for Microsoft 365 recognizes these variables as authentication defaults.
      if (tenantId) env.CLIMICROSOFT365_TENANT = tenantId;
      if (clientId) env.CLIMICROSOFT365_ENTRAAPPID = clientId;

      // Plugin-owned convenience variables for skills/bootstrap prompts and shell snippets.
      if (tenantId) env.M365_TENANT_ID = tenantId;
      if (clientId) env.M365_CLIENT_ID = clientId;
      if (account) env.M365_ACCOUNT = account;
      if (connectionName) env.M365_CONNECTION_NAME = connectionName;

      return env;
    }, { priority: 20 });
  }
});
