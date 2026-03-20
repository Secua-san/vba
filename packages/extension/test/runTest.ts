import path from "node:path";
import process from "node:process";
import { runTests } from "@vscode/test-electron";

async function main(): Promise<void> {
  const extensionDevelopmentPath = path.resolve(__dirname, "..", "..");
  const extensionTestsPath = path.resolve(__dirname, "suite", "index.js");
  const fixtureWorkspace = path.resolve(extensionDevelopmentPath, "test", "fixtures");

  try {
    await runTests({
      extensionDevelopmentPath,
      extensionTestsPath,
      launchArgs: [fixtureWorkspace, "--disable-extensions"]
    });
  } catch (error) {
    console.error(error);
    process.exitCode = 1;
  }
}

void main();
