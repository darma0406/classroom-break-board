#!/usr/bin/env node

const fs = require("fs");
const path = require("path");

const ROOT = process.cwd();
const IGNORED_DIRS = new Set([
  ".git",
  "node_modules",
  "scripts",
  ".tmp-chrome-profile",
  ".tmp-chrome-profile-2",
  ".tmp-chrome-profile-legacy",
]);

const GLOBALS = new Set([
  "Array",
  "Boolean",
  "Date",
  "Error",
  "JSON",
  "Math",
  "Number",
  "Object",
  "Promise",
  "RegExp",
  "Set",
  "String",
  "URL",
  "URLSearchParams",
  "clearInterval",
  "clearTimeout",
  "console",
  "document",
  "fetch",
  "firebase",
  "history",
  "localStorage",
  "location",
  "setInterval",
  "setTimeout",
  "window",
]);

const KEYWORDS = new Set([
  "async",
  "await",
  "break",
  "case",
  "catch",
  "const",
  "continue",
  "default",
  "delete",
  "do",
  "else",
  "false",
  "finally",
  "for",
  "function",
  "if",
  "in",
  "instanceof",
  "let",
  "new",
  "null",
  "of",
  "return",
  "switch",
  "throw",
  "true",
  "try",
  "typeof",
  "undefined",
  "var",
  "while",
]);

function findIndexHtml(dir) {
  const entries = fs.readdirSync(dir, { withFileTypes: true });

  for (const entry of entries) {
    if (entry.isFile() && entry.name === "index.html") {
      return path.join(dir, entry.name);
    }
  }

  for (const entry of entries) {
    if (!entry.isDirectory() || IGNORED_DIRS.has(entry.name)) continue;
    const found = findIndexHtml(path.join(dir, entry.name));
    if (found) return found;
  }

  return "";
}

function stripStringsAndComments(source) {
  return source
    .replace(/\/\*[\s\S]*?\*\//g, " ")
    .replace(/\/\/[^\n\r]*/g, " ");
}

function extractInlineScripts(html) {
  return [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi)]
    .map((match) => match[1])
    .filter((script) => script.trim());
}

function collectDefinitions(source) {
  const definitions = new Set();
  const patterns = [
    /\b(?:async\s+)?function\s+([A-Za-z_$][\w$]*)\s*\(/g,
    /\b(?:const|let|var)\s+([A-Za-z_$][\w$]*)\s*=\s*(?:async\s*)?(?:function\b|\([^)]*\)\s*=>|[A-Za-z_$][\w$]*\s*=>)/g,
  ];

  for (const pattern of patterns) {
    for (const match of source.matchAll(pattern)) {
      definitions.add(match[1]);
    }
  }

  return definitions;
}

function addNamesFromList(bindings, text) {
  for (const item of text.split(",")) {
    const match = item.trim().match(/^([A-Za-z_$][\w$]*)$/);
    if (match) bindings.add(match[1]);
  }
}

function collectBindings(source, definitions) {
  const bindings = new Set(definitions);
  const patterns = [
    /\b(?:const|let|var)\s+([A-Za-z_$][\w$]*)\b/g,
    /\b(?:async\s+)?function\s*[A-Za-z_$]*\s*\(([^)]*)\)/g,
    /\bcatch\s*\(([^)]*)\)/g,
  ];

  for (const pattern of patterns) {
    for (const match of source.matchAll(pattern)) {
      addNamesFromList(bindings, match[1]);
    }
  }

  for (const match of source.matchAll(/\(([^)]*)\)\s*=>/g)) {
    addNamesFromList(bindings, match[1]);
  }

  for (const match of source.matchAll(/[=(:,]\s*\(([^()]*)\)\s*=>/g)) {
    addNamesFromList(bindings, match[1]);
  }

  for (const match of source.matchAll(/\b([A-Za-z_$][\w$]*)\s*=>/g)) {
    bindings.add(match[1]);
  }

  return bindings;
}

function isPropertyCall(source, index) {
  let i = index - 1;
  while (i >= 0 && /\s/.test(source[i])) i--;
  return source[i] === "." || source[i] === "?";
}

function collectCalls(source) {
  const calls = new Set();
  const callPattern = /\b([A-Za-z_$][\w$]*)\s*\(/g;

  for (const match of source.matchAll(callPattern)) {
    const name = match[1];
    if (KEYWORDS.has(name) || GLOBALS.has(name)) continue;
    if (isPropertyCall(source, match.index)) continue;
    calls.add(name);
  }

  return calls;
}

function main() {
  const indexPath = findIndexHtml(ROOT);
  if (!indexPath) {
    console.error("FAIL: index.html not found");
    process.exit(1);
  }

  const html = fs.readFileSync(indexPath, "utf8");
  const scripts = extractInlineScripts(html);
  const source = stripStringsAndComments(scripts.join("\n"));
  const definitions = collectDefinitions(source);
  const bindings = collectBindings(source, definitions);
  const calls = collectCalls(source);
  const missing = [...calls]
    .filter((name) => !bindings.has(name))
    .sort((a, b) => a.localeCompare(b));

  if (!missing.length) {
    console.log("PASS: no missing functions");
    return;
  }

  console.log("Missing functions:");
  for (const name of missing) {
    console.log(`- ${name}`);
  }
  process.exitCode = 1;
}

main();
