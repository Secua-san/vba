import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { signatureMemberAllowList } from "../lib/referenceSignatureConfig.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const referenceFile = path.resolve(__dirname, "..", "..", "resources", "reference", "mslearn-vba-reference.json");
const auditedMembers = new Map(
  [...signatureMemberAllowList]
    .filter(([ownerName]) => ownerName !== "Application")
    .map(([ownerName, memberNames]) => [ownerName, new Set(memberNames)]),
);

async function loadReferenceData() {
  return JSON.parse(await readFile(referenceFile, "utf8"));
}

function getSignature(data, ownerName, memberName) {
  const owner = data.excel.objectModel.items.find((item) => item.name === ownerName);
  const member = owner?.sections.flatMap((section) => section.members).find((entry) => entry.name === memberName);
  return member?.signature;
}

function hasSequentialNumericSuffixParameters(parameters) {
  if (parameters.length < 3) {
    return false;
  }

  const matches = parameters.map((parameter) => parameter.name.match(/^([A-Za-z_]+)(\d+)$/u));

  if (matches.some((match) => !match?.[1] || !match[2])) {
    return false;
  }

  const prefix = matches[0][1].toLowerCase();
  return matches.every((match, index) => match[1].toLowerCase() === prefix && Number(match[2]) === index + 1);
}

function hasUniformVariadicCountDescription(parameters) {
  const firstDescription = parameters[0]?.description?.trim();

  if (!firstDescription) {
    return false;
  }

  const normalizedDescriptions = parameters.map((parameter) => parameter.description?.trim() ?? "");

  if (!normalizedDescriptions.every((description) => description === firstDescription)) {
    return false;
  }

  return /\b1\s*(?:to|-)\s*\d+\b/iu.test(firstDescription) || /\bbetween\s+1\s+and\s+\d+\b/iu.test(firstDescription);
}

test("監査対象の署名データは引数メタデータ欠落なく保持される", async () => {
  const data = await loadReferenceData();

  for (const [ownerName, memberNames] of auditedMembers) {
    for (const memberName of memberNames) {
      const signature = getSignature(data, ownerName, memberName);

      assert.ok(signature, `${ownerName}.${memberName} の署名が必要です`);
      assert.ok(signature.returnType, `${ownerName}.${memberName} の戻り値型が必要です`);
      assert.ok(signature.parameters.length > 0, `${ownerName}.${memberName} の引数が必要です`);

      for (const parameter of signature.parameters) {
        assert.ok(parameter.dataType, `${ownerName}.${memberName}.${parameter.name} の型が必要です`);
        assert.ok(parameter.description, `${ownerName}.${memberName}.${parameter.name} の説明が必要です`);
        assert.notEqual(
          parameter.isRequired,
          undefined,
          `${ownerName}.${memberName}.${parameter.name} の必須/省略可能フラグが必要です`,
        );
        assert.ok(
          parameter.label.includes(" As "),
          `${ownerName}.${memberName}.${parameter.name} のラベルに型情報が必要です`,
        );
      }

      if (
        hasSequentialNumericSuffixParameters(signature.parameters) &&
        signature.label.includes("...") &&
        hasUniformVariadicCountDescription(signature.parameters)
      ) {
        assert.equal(
          signature.parameters[0]?.isRequired,
          true,
          `${ownerName}.${memberName} の第1引数は必須である必要があります`,
        );

        for (const parameter of signature.parameters.slice(1)) {
          assert.equal(
            parameter.isRequired,
            false,
            `${ownerName}.${memberName}.${parameter.name} は省略可能である必要があります`,
          );
        }
      }
    }
  }
});

test("WorksheetFunction の現行 Learn スナップショットに XLookup / XMATCH は未掲載", async () => {
  const data = await loadReferenceData();
  const worksheetFunction = data.excel.objectModel.items.find((item) => item.name === "WorksheetFunction");
  const memberNames = new Set(
    worksheetFunction?.sections.flatMap((section) => section.members).map((member) => member.name.toLowerCase()) ?? [],
  );

  assert.equal(memberNames.has("xlookup"), false);
  assert.equal(memberNames.has("xmatch"), false);
});
