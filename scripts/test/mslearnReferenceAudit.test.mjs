import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { signatureMemberAllowList, signatureMissingMemberWatchList } from "../lib/referenceSignatureConfig.mjs";

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

function getOwnerMemberNames(data, ownerName) {
  const owner = data.excel.objectModel.items.find((item) => item.name === ownerName);
  const memberNames = new Set(
    owner?.sections.flatMap((section) => section.members).map((member) => String(member.name).toLowerCase()) ?? [],
  );

  return {
    memberNames,
    owner,
  };
}

function buildMissingMemberGuidance(ownerName, memberName) {
  return `${ownerName}.${memberName} が Learn スナップショットへ追加されました。scripts/lib/referenceSignatureConfig.mjs の watch list から外し、allow list / 再生成 / server・extension テストの更新を docs/process/mslearn-signature-regeneration.md に従って進めてください。`;
}

function normalizeMemberName(memberName) {
  return String(memberName).toLowerCase();
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

test("未掲載監視対象は owner 単位の watch list で管理される", async () => {
  const data = await loadReferenceData();

  for (const [ownerName, memberNamesToWatch] of signatureMissingMemberWatchList) {
    const { owner, memberNames } = getOwnerMemberNames(data, ownerName);

    assert.ok(owner, `${ownerName} owner が Learn スナップショット内に必要です`);

    for (const memberName of memberNamesToWatch) {
      assert.equal(memberNames.has(normalizeMemberName(memberName)), false, buildMissingMemberGuidance(ownerName, memberName));
    }
  }
});

test("未掲載監視対象は allow list と重複しない", () => {
  for (const [ownerName, memberNamesToWatch] of signatureMissingMemberWatchList) {
    const normalizedAllowListedMembers = new Set(
      [...(signatureMemberAllowList.get(ownerName) ?? new Set())].map((memberName) => normalizeMemberName(memberName)),
    );

    for (const memberName of memberNamesToWatch) {
      assert.equal(
        normalizedAllowListedMembers.has(normalizeMemberName(memberName)),
        false,
        `${ownerName}.${memberName} は watch list と allow list に同時登録できません`,
      );
    }
  }
});

test("未掲載監視対象は owner 内で大文字小文字違いの重複を持たない", () => {
  for (const [ownerName, memberNamesToWatch] of signatureMissingMemberWatchList) {
    const normalizedMemberNames = [...memberNamesToWatch].map((memberName) => normalizeMemberName(memberName));

    assert.equal(
      new Set(normalizedMemberNames).size,
      normalizedMemberNames.length,
      `${ownerName} の watch list に大文字小文字違いの重複があります`,
    );
  }
});
