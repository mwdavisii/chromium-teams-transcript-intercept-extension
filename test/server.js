// test/server.js
// Minimal HTTPS fixture server for the integration test.
//
// Serves three resources over a self-signed certificate so the page's own
// fetch() (running under a `*.sharepoint.com` origin) can reach a metadata
// endpoint whose response the extension's MAIN-world hook observes:
//
//   GET /stream.html                              -> test/fixtures/stream.html
//   GET /_api/v2.1/drives/me/items/transcripts    -> test/fixtures/metadata.json
//   GET /transcript-content                       -> test/fixtures/transcript.json
//
// The metadata JSON's `temporaryDownloadUrl` points at `/transcript-content`,
// so the content-script fetch pipeline also stays on the same origin.
//
// HTTPS + a self-signed cert is used because the extension's content scripts
// only match `https://*.sharepoint.com/*`. The integration test launches
// Chromium with
//   --host-resolver-rules="MAP fakefixture.sharepoint.com 127.0.0.1"
//   --ignore-certificate-errors
// and navigates to https://fakefixture.sharepoint.com:9876/stream.html, so
// this server never actually needs a `*.sharepoint.com` cert (the browser
// trusts it via --ignore-certificate-errors; the hostname rewrite is done
// purely in Chromium's resolver). The cert's SAN still lists the fixture
// hostname for environments that validate it.
//
// Kept as a plain Node `https` server with no framework, per the task spec.

'use strict';

const https = require('https');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const FIXTURES_DIR = path.join(__dirname, 'fixtures');
const CERTS_DIR = path.join(__dirname, 'fixtures', 'certs');
const CERT_PATH = path.join(CERTS_DIR, 'cert.pem');
const KEY_PATH = path.join(CERTS_DIR, 'key.pem');

const HOSTNAME = 'fakefixture.sharepoint.com';
const PORT = 9876;

// ---------------------------------------------------------------------------
// Self-signed cert generation (idempotent — only when missing).
// ---------------------------------------------------------------------------
function ensureCertificate() {
  if (fs.existsSync(CERT_PATH) && fs.existsSync(KEY_PATH)) {
    return { key: fs.readFileSync(KEY_PATH), cert: fs.readFileSync(CERT_PATH) };
  }
  fs.mkdirSync(CERTS_DIR, { recursive: true });
  try {
    execFileSync('openssl', [
      'req', '-x509', '-newkey', 'rsa:2048',
      '-keyout', KEY_PATH,
      '-out', CERT_PATH,
      '-days', '3650',
      '-nodes',
      '-subj', '/CN=' + HOSTNAME,
      '-addext', 'subjectAltName=DNS:' + HOSTNAME + ',DNS:*.' + HOSTNAME.split('.').slice(1).join('.'),
    ], { stdio: 'ignore' });
  } catch (e) {
    throw new Error(
      'openssl failed to generate the fixture certificate — is openssl installed? ' +
      String((e && e.message) || e)
    );
  }
  return { key: fs.readFileSync(KEY_PATH), cert: fs.readFileSync(CERT_PATH) };
}

// ---------------------------------------------------------------------------
// Route table.
// ---------------------------------------------------------------------------
const ROUTES = {
  '/stream.html': {
    file: path.join(FIXTURES_DIR, 'stream.html'),
    type: 'text/html; charset=utf-8',
  },
  '/_api/v2.1/drives/me/items/transcripts': {
    file: path.join(FIXTURES_DIR, 'metadata.json'),
    type: 'application/json; charset=utf-8',
  },
  '/transcript-content': {
    file: path.join(FIXTURES_DIR, 'transcript.json'),
    type: 'application/json; charset=utf-8',
  },
};

// Permissive CORS so the page fetch succeeds from any origin. The fixture page
// is same-origin with these endpoints anyway, but keep it permissive so the
// content-script fetch (which may carry a bearer header) never preflights into
// failure under Chromium's stricter enforcement.
const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
  'Access-Control-Allow-Headers': '*',
  'Access-Control-Allow-Credentials': 'true',
};

let server = null;

function startServer() {
  const { key, cert } = ensureCertificate();

  server = https.createServer({ key, cert }, (req, res) => {
    const url = (req.url || '').split('?')[0];

    if (req.method === 'OPTIONS') {
      res.writeHead(204, CORS_HEADERS);
      res.end();
      return;
    }

    const route = ROUTES[url];
    if (!route) {
      res.writeHead(404, { ...CORS_HEADERS, 'Content-Type': 'text/plain' });
      res.end('not found');
      return;
    }

    try {
      const body = fs.readFileSync(route.file);
      res.writeHead(200, { ...CORS_HEADERS, 'Content-Type': route.type });
      res.end(body);
    } catch (e) {
      res.writeHead(500, { ...CORS_HEADERS, 'Content-Type': 'text/plain' });
      res.end('fixture read error: ' + String((e && e.message) || e));
    }
  });

  return new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(PORT, '127.0.0.1', () => {
      resolve();
    });
  });
}

function stopServer() {
  return new Promise((resolve) => {
    if (!server) {
      resolve();
      return;
    }
    server.close(() => {
      server = null;
      resolve();
    });
    // Force-close any keep-alive sockets that would otherwise delay close().
    server.closeAllConnections && server.closeAllConnections();
  });
}

module.exports = { startServer, stopServer, PORT, HOSTNAME };
