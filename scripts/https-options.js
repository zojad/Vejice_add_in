const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFileSync } = require("child_process");

const OFFICE_CERT_DIR = path.join(os.homedir(), ".office-addin-dev-certs");
const FALLBACK_CERT_DIR = path.join(os.homedir(), ".vejice-dev-certs");

function readHttpsOptions({ caPath, certPath, keyPath }) {
  return {
    ca: fs.readFileSync(caPath),
    cert: fs.readFileSync(certPath),
    key: fs.readFileSync(keyPath),
  };
}

function hasFiles(paths) {
  return paths.every((filePath) => typeof filePath === "string" && fs.existsSync(filePath));
}

function resolveExplicitHttpsOptions() {
  const keyPath = process.env.VEJICE_DEV_SSL_KEY_PATH || "";
  const certPath = process.env.VEJICE_DEV_SSL_CERT_PATH || "";
  const caPath = process.env.VEJICE_DEV_SSL_CA_PATH || certPath;
  if (!keyPath || !certPath) return null;
  if (!hasFiles([keyPath, certPath, caPath])) {
    throw new Error("Configured VEJICE_DEV_SSL_* certificate files were not found.");
  }
  return {
    source: "env",
    options: readHttpsOptions({ caPath, certPath, keyPath }),
  };
}

function resolveOfficeAddinHttpsOptions() {
  const caPath = path.join(OFFICE_CERT_DIR, "ca.crt");
  const certPath = path.join(OFFICE_CERT_DIR, "localhost.crt");
  const keyPath = path.join(OFFICE_CERT_DIR, "localhost.key");
  if (!hasFiles([caPath, certPath, keyPath])) return null;
  return {
    source: "office-addin-dev-certs",
    options: readHttpsOptions({ caPath, certPath, keyPath }),
  };
}

function ensureFallbackCertificateFiles() {
  fs.mkdirSync(FALLBACK_CERT_DIR, { recursive: true });
  const keyPath = path.join(FALLBACK_CERT_DIR, "localhost.key");
  const certPath = path.join(FALLBACK_CERT_DIR, "localhost.crt");
  const configPath = path.join(FALLBACK_CERT_DIR, "openssl.cnf");

  if (hasFiles([keyPath, certPath])) {
    return { keyPath, certPath };
  }

  const config = `
[req]
distinguished_name = req_distinguished_name
x509_extensions = v3_req
prompt = no

[req_distinguished_name]
CN = localhost

[v3_req]
subjectAltName = @alt_names
keyUsage = digitalSignature, keyEncipherment
extendedKeyUsage = serverAuth

[alt_names]
DNS.1 = localhost
IP.1 = 127.0.0.1
IP.2 = ::1
`.trim();

  fs.writeFileSync(configPath, `${config}\n`, "utf8");
  try {
    execFileSync(
      "openssl",
      [
        "req",
        "-x509",
        "-nodes",
        "-newkey",
        "rsa:2048",
        "-keyout",
        keyPath,
        "-out",
        certPath,
        "-days",
        "30",
        "-config",
        configPath,
        "-extensions",
        "v3_req",
      ],
      { stdio: "ignore" }
    );
  } catch (error) {
    throw new Error(
      `Unable to generate fallback HTTPS certificate with openssl. ${error.message || error}`
    );
  }

  return { keyPath, certPath };
}

function resolveFallbackHttpsOptions() {
  const { keyPath, certPath } = ensureFallbackCertificateFiles();
  return {
    source: "generated-self-signed",
    options: readHttpsOptions({ caPath: certPath, certPath, keyPath }),
  };
}

function getLocalHttpsOptions() {
  const explicit = resolveExplicitHttpsOptions();
  if (explicit) return explicit;

  const officeAddin = resolveOfficeAddinHttpsOptions();
  if (officeAddin) return officeAddin;

  return resolveFallbackHttpsOptions();
}

module.exports = {
  getLocalHttpsOptions,
};
