"use strict";

const fs = require("node:fs");
const path = require("node:path");
const { spawnSync } = require("node:child_process");

const root = path.join(__dirname, "..");
const syntaxOnly = process.argv.includes("--syntax-only");

function walk(directory, predicate) {
  if (!fs.existsSync(directory)) return [];
  return fs.readdirSync(directory, { withFileTypes: true })
    .flatMap(entry => {
      const fullPath = path.join(directory, entry.name);
      if (entry.isDirectory()) return walk(fullPath, predicate);
      return predicate(fullPath) ? [fullPath] : [];
    })
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
}

function runNode(args, label) {
  const result = spawnSync(process.execPath, args, {
    cwd: root,
    stdio: "inherit",
    env: process.env
  });
  if (result.error) {
    console.error(`${label}: ${result.error.message}`);
    return false;
  }
  if (result.status !== 0) {
    console.error(`${label}: failed with exit code ${result.status}`);
    return false;
  }
  return true;
}

const sourceFiles = walk(path.join(root, "js"), filename => filename.endsWith(".js"));
let failed = false;

console.log(`Checking JavaScript syntax (${sourceFiles.length} files)...`);
sourceFiles.forEach(filename => {
  if (!runNode(["--check", filename], path.relative(root, filename))) failed = true;
});

const jsonFiles = [
  path.join(root, "package.json"),
  ...walk(path.join(root, "schemas"), filename => filename.endsWith(".json")),
  ...walk(path.join(root, "tests", "fixtures"), filename => filename.endsWith(".json")),
  ...walk(path.join(root, "tests", "expected"), filename => filename.endsWith(".json"))
];
console.log(`Checking JSON syntax (${jsonFiles.length} files)...`);
jsonFiles.forEach(filename => {
  try {
    JSON.parse(fs.readFileSync(filename, "utf8"));
  } catch (error) {
    failed = true;
    console.error(`${path.relative(root, filename)}: ${error.message}`);
  }
});

if (!syntaxOnly) {
  const testFiles = walk(__dirname, filename => filename.endsWith(".test.js"));
  console.log(`Running regression tests (${testFiles.length} files)...`);
  testFiles.forEach(filename => {
    console.log(`\n> ${path.relative(root, filename)}`);
    if (!runNode([filename], path.relative(root, filename))) failed = true;
  });
}

if (failed) {
  console.error("\nOne or more checks failed.");
  process.exit(1);
}

console.log(syntaxOnly ? "\nAll syntax checks passed." : "\nAll tests passed.");
