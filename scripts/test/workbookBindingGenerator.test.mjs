import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import test from "node:test";

const execFileAsync = promisify(execFile);
const generatorPath = path.resolve("scripts", "generate-workbook-binding.mjs");

test("package は workbook binding generator の npm script を公開する", async () => {
  const packageManifest = JSON.parse(await readFile(path.resolve("package.json"), "utf8"));

  assert.equal(packageManifest.scripts["generate:workbook-binding"], "node scripts/generate-workbook-binding.mjs");
});

test("CLI は workbook binding manifest を bundle root の正本パスへ書き出す", async () => {
  const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "vba-workbook-binding-"));
  const workbookPath = path.join(temporaryDirectory, "Book1.xlsm");
  const bundleRoot = path.join(temporaryDirectory, "bundle");
  const manifestPath = path.join(bundleRoot, ".vba", "workbook-binding.json");

  try {
    await writeFile(workbookPath, "");

    await execFileAsync(process.execPath, [
      generatorPath,
      "--workbook-path",
      workbookPath,
      "--bundle-root",
      bundleRoot,
    ], {
      cwd: path.resolve("."),
    });

    const manifest = JSON.parse(await readFile(manifestPath, "utf8"));

    assert.deepEqual(manifest, {
      version: 1,
      artifact: "workbook-binding-manifest",
      bindingKind: "active-workbook-fullname",
      workbook: {
        fullName: path.resolve(workbookPath),
        name: "Book1.xlsm",
        path: path.dirname(path.resolve(workbookPath)),
        isAddIn: false,
        sourceKind: "openxml-package",
      },
    });
  } finally {
    await rm(temporaryDirectory, { force: true, recursive: true });
  }
});

test("CLI は .xla / .xlam workbook を add-in として拒否する", async () => {
  const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "vba-workbook-binding-"));

  try {
    for (const extension of [".xla", ".xlam"]) {
      const workbookPath = path.join(temporaryDirectory, `Addin${extension}`);
      const bundleRoot = path.join(temporaryDirectory, `bundle-${extension.slice(1)}`);

      await writeFile(workbookPath, "");

      await assert.rejects(
        execFileAsync(process.execPath, [
          generatorPath,
          "--workbook-path",
          workbookPath,
          "--bundle-root",
          bundleRoot,
        ], {
          cwd: path.resolve("."),
        }),
        /add-in/u,
      );

      await assert.rejects(stat(path.join(bundleRoot, ".vba", "workbook-binding.json")));
    }
  } finally {
    await rm(temporaryDirectory, { force: true, recursive: true });
  }
});

test("CLI は --workbook-path と --bundle-root 以外の入力を拒否する", async () => {
  const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "vba-workbook-binding-"));
  const workbookPath = path.join(temporaryDirectory, "Book1.xlsm");
  const bundleRoot = path.join(temporaryDirectory, "bundle");

  try {
    await writeFile(workbookPath, "");

    await assert.rejects(
      execFileAsync(process.execPath, [
        generatorPath,
        "--workbook-path",
        workbookPath,
        "--bundle-root",
        bundleRoot,
        "--out",
        path.join(temporaryDirectory, "binding.json"),
      ], {
        cwd: path.resolve("."),
      }),
      /未対応の引数/u,
    );
  } finally {
    await rm(temporaryDirectory, { force: true, recursive: true });
  }
});

test("CLI は option の値に別の flag が来た場合に値不足として拒否する", async () => {
  const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "vba-workbook-binding-"));
  const workbookPath = path.join(temporaryDirectory, "Book1.xlsm");

  try {
    await writeFile(workbookPath, "");

    await assert.rejects(
      execFileAsync(process.execPath, [
        generatorPath,
        "--bundle-root",
        "--workbook-path",
        workbookPath,
      ], {
        cwd: path.resolve("."),
      }),
      /--bundle-root の後に bundle root パスが必要です/u,
    );

    await assert.rejects(
      execFileAsync(process.execPath, [
        generatorPath,
        "--workbook-path",
        "--bundle-root",
        temporaryDirectory,
      ], {
        cwd: path.resolve("."),
      }),
      /--workbook-path の後に workbook パスが必要です/u,
    );
  } finally {
    await rm(temporaryDirectory, { force: true, recursive: true });
  }
});
