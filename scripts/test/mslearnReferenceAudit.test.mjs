import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { signatureMemberAllowList, signatureMissingMemberWatchList } from "../lib/referenceSignatureConfig.mjs";
import {
  dialogFrameMethodMemberNames,
  dialogFramePropertyMemberNames,
  dialogSheetCommonCallableMemberNames,
  dialogSheetPropertyMemberNames,
  supplementalOwnerMemberOverrides,
} from "../lib/supplementalReferenceConfig.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const referenceFile = path.resolve(__dirname, "..", "..", "resources", "reference", "mslearn-vba-reference.json");
const auditedMembers = new Map(
  [...signatureMemberAllowList]
    .filter(([ownerName]) => ownerName !== "Application")
    .map(([ownerName, memberNames]) => [ownerName, new Set(memberNames)]),
);
const expectedDialogSheetControlCollectionMemberNames = ["Buttons", "CheckBoxes", "OptionButtons"];
const expectedDialogSheetControlCollectionOwnerConfigs = [
  {
    collectionName: "Buttons",
    itemMethodMemberNames: ["Select"],
    itemName: "Button",
    itemPropertyMemberNames: ["Caption", "Name", "OnAction", "Text"],
  },
  {
    collectionName: "CheckBoxes",
    itemMethodMemberNames: ["Select"],
    itemName: "CheckBox",
    itemPropertyMemberNames: ["Caption", "Name", "OnAction", "Text", "Value"],
  },
  {
    collectionName: "OptionButtons",
    itemMethodMemberNames: ["Select"],
    itemName: "OptionButton",
    itemPropertyMemberNames: ["Caption", "Name", "OnAction", "Text", "Value"],
  },
];

async function loadReferenceData() {
  return JSON.parse(await readFile(referenceFile, "utf8"));
}

function getSignature(data, ownerName, memberName) {
  const member = getMember(data, ownerName, memberName);
  return member?.signature;
}

function getMember(data, ownerName, memberName) {
  const owner = data.excel.objectModel.items.find((item) => item.name === ownerName);
  return owner?.sections.flatMap((section) => section.members).find((entry) => entry.name === memberName);
}

function getOwnerMemberNames(data, ownerName) {
  const owner = data.excel.objectModel.items.find((item) => item.name === ownerName);
  const memberNames = new Set(
    owner?.sections.flatMap((section) => section.members).map((member) => normalizeMemberName(member.name)) ?? [],
  );

  return {
    memberNames,
    owner,
  };
}

function buildMissingMemberGuidance(ownerName, memberName) {
  return `${ownerName}.${memberName} が Learn スナップショットへ追加されました。scripts/lib/referenceSignatureConfig.mjs の watch list から外し、docs/process/mslearn-signature-regeneration.md の手順に従って更新してください。`;
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

function hasCollapsedSentenceBoundary(value) {
  const sanitizedValue = value.replace(/\b[A-Za-z][A-Za-z0-9_]*\.[A-Za-z][A-Za-z0-9_]*\b/gu, "ApiReference");
  return /([a-z0-9)])\.(?=(?:[A-Z][a-z]|[A-Z]{2,}:))/u.test(sanitizedValue);
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
        assert.equal(
          hasCollapsedSentenceBoundary(parameter.description),
          false,
          `${ownerName}.${memberName}.${parameter.name} の説明文で文境界の空白が欠落しています`,
        );
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

test("DialogSheet 補助ソースは allow list の property / method だけを保持する", async () => {
  const data = await loadReferenceData();
  const { owner, memberNames } = getOwnerMemberNames(data, "DialogSheet");
  const normalizedAllowedMemberNames = new Set(
    [...dialogSheetPropertyMemberNames, ...dialogSheetCommonCallableMemberNames, ...expectedDialogSheetControlCollectionMemberNames].map(
      (memberName) => normalizeMemberName(memberName),
    ),
  );

  assert.ok(owner, "DialogSheet owner が補助ソースとして必要です");
  assert.deepEqual([...memberNames].sort(), [...normalizedAllowedMemberNames].sort());
});

test("DialogSheet 補助ソースは allow list 全件の署名を完全に保持する", async () => {
  const data = await loadReferenceData();

  for (const memberName of dialogSheetCommonCallableMemberNames) {
    const signature = getSignature(data, "DialogSheet", memberName);

    assert.ok(signature, `DialogSheet.${memberName} の署名が必要です`);
    assert.equal(signature.ownerName, "DialogSheet", `DialogSheet.${memberName} の ownerName が必要です`);
  }

  assert.equal(
    getSignature(data, "DialogSheet", "Activate")?.parameters.length,
    0,
    "DialogSheet.Activate は引数なしの署名である必要があります",
  );
  assert.equal(
    getSignature(data, "DialogSheet", "Evaluate")?.parameters.length,
    1,
    "DialogSheet.Evaluate は単一引数の署名である必要があります",
  );
  assert.equal(
    getSignature(data, "DialogSheet", "ExportAsFixedFormat")?.parameters.length,
    9,
    "DialogSheet.ExportAsFixedFormat は 9 引数の署名である必要があります",
  );
  assert.equal(
    getSignature(data, "DialogSheet", "SaveAs")?.parameters.length,
    10,
    "DialogSheet.SaveAs は 10 引数の署名である必要があります",
  );

  for (const memberName of expectedDialogSheetControlCollectionMemberNames) {
    const member = getMember(data, "DialogSheet", memberName);
    const signature = getSignature(data, "DialogSheet", memberName);

    assert.ok(member, `DialogSheet.${memberName} member が必要です`);
    assert.equal(member.typeName, memberName, `DialogSheet.${memberName} は ${memberName} owner へ接続する必要があります`);
    assert.ok(signature, `DialogSheet.${memberName} の署名が必要です`);
    assert.equal(signature.ownerName, "DialogSheet", `DialogSheet.${memberName} の ownerName が必要です`);
    assert.equal(signature.parameters.length, 1, `DialogSheet.${memberName} は単一 selector の署名である必要があります`);
  }
});

test("DialogSheet control collection surface は想定 3 member のみを公開する", async () => {
  const data = await loadReferenceData();
  const dialogSheetOwner = data.excel.objectModel.items.find((item) => item.name === "DialogSheet");
  const controlCollectionOwnerNames = new Set(
    expectedDialogSheetControlCollectionOwnerConfigs.map((config) => normalizeMemberName(config.collectionName)),
  );
  const controlCollectionMembers =
    dialogSheetOwner?.sections
      .flatMap((section) => section.members)
      .filter((member) => controlCollectionOwnerNames.has(normalizeMemberName(member.typeName ?? "")))
      .map((member) => normalizeMemberName(member.name)) ?? [];

  assert.deepEqual(
    [...new Set(controlCollectionMembers)].sort(),
    expectedDialogSheetControlCollectionMemberNames.map((memberName) => normalizeMemberName(memberName)).sort(),
    "DialogSheet は Buttons / CheckBoxes / OptionButtons だけを control collection surface として公開する必要があります",
  );
});

test("DialogSheet 補助 property は DialogFrame owner へ型付けする", async () => {
  const data = await loadReferenceData();
  const dialogFrameMember = getMember(data, "DialogSheet", "DialogFrame");

  assert.ok(dialogFrameMember, "DialogSheet.DialogFrame member が必要です");
  assert.equal(dialogFrameMember.typeName, "DialogFrame");
  assert.equal(dialogFrameMember.signature, undefined);
});

test("DialogSheet 補助ソースは dummy / legacy member と正規化重複を含まない", async () => {
  const data = await loadReferenceData();
  const { owner } = getOwnerMemberNames(data, "DialogSheet");
  const rawMemberNames = owner?.sections.flatMap((section) => section.members).map((member) => String(member.name)) ?? [];
  const normalizedMemberNames = rawMemberNames.map((memberName) => normalizeMemberName(memberName));

  assert.ok(owner, "DialogSheet owner が補助ソースとして必要です");

  for (const memberName of rawMemberNames) {
    assert.equal(/^_dummy/iu.test(memberName), false, `DialogSheet.${memberName} は dummy member を含められません`);
    assert.equal(/^_/u.test(memberName), false, `DialogSheet.${memberName} は legacy member を含められません`);
  }

  assert.equal(
    new Set(normalizedMemberNames).size,
    normalizedMemberNames.length,
    "DialogSheet owner に正規化後の重複 member 名があります",
  );
});

test("DialogSheets clone は Item を DialogSheet として型付けする", async () => {
  const data = await loadReferenceData();
  const itemMember = getMember(data, "DialogSheets", "Item");

  assert.ok(itemMember, "DialogSheets.Item member が必要です");
  assert.equal(itemMember.typeName, "DialogSheet");
});

test("OLEObjects member override は Item を OLEObjects として型付けする", async () => {
  const data = await loadReferenceData();
  const itemMember = getMember(data, "OLEObjects", "Item");
  const itemOverride = supplementalOwnerMemberOverrides
    .flatMap((ownerConfig) => ownerConfig.members.map((member) => ({ member, ownerName: ownerConfig.ownerName })))
    .find((entry) => entry.ownerName === "OLEObjects" && entry.member.name === "Item");

  assert.ok(itemMember, "OLEObjects.Item member が必要です");
  assert.ok(itemOverride, "OLEObjects.Item override 設定が必要です");
  assert.equal(itemMember.typeName, "OLEObjects");
});

test("DialogSheet 補助 root は Application / Workbook.DialogSheets を DialogSheets として型付けする", async () => {
  const data = await loadReferenceData();

  for (const ownerName of ["Application", "Workbook"]) {
    const member = getMember(data, ownerName, "DialogSheets");

    assert.ok(member, `${ownerName}.DialogSheets member が必要です`);
    assert.equal(member.typeName, "DialogSheets", `${ownerName}.DialogSheets は DialogSheets owner へ接続する必要があります`);
  }
});

test("DialogFrame 補助ソースは allow list の property / method だけを保持する", async () => {
  const data = await loadReferenceData();
  const { owner, memberNames } = getOwnerMemberNames(data, "DialogFrame");
  const normalizedAllowedMemberNames = new Set(
    [...dialogFramePropertyMemberNames, ...dialogFrameMethodMemberNames].map((memberName) =>
      normalizeMemberName(memberName),
    ),
  );

  assert.ok(owner, "DialogFrame owner が補助ソースとして必要です");
  assert.deepEqual([...memberNames].sort(), [...normalizedAllowedMemberNames].sort());
});

test("DialogFrame 補助 property / method は型情報と署名を保持する", async () => {
  const data = await loadReferenceData();

  for (const memberName of dialogFramePropertyMemberNames) {
    const member = getMember(data, "DialogFrame", memberName);

    assert.ok(member, `DialogFrame.${memberName} member が必要です`);
    assert.equal(member.typeName, "String", `DialogFrame.${memberName} は String 型である必要があります`);
    assert.equal(member.signature, undefined, `DialogFrame.${memberName} は property として扱う必要があります`);
  }

  const selectSignature = getSignature(data, "DialogFrame", "Select");

  assert.ok(selectSignature, "DialogFrame.Select の署名が必要です");
  assert.equal(selectSignature.ownerName, "DialogFrame");
  assert.equal(selectSignature.label, "Select(Replace) As Object");
  assert.equal(selectSignature.parameters.length, 1);
  assert.equal(selectSignature.parameters[0]?.name, "Replace");
  assert.equal(selectSignature.parameters[0]?.dataType, "Object");
  assert.equal(selectSignature.parameters[0]?.isRequired, false);
});

test("DialogFrame 補助ソースは dummy / legacy member と正規化重複を含まない", async () => {
  const data = await loadReferenceData();
  const { owner } = getOwnerMemberNames(data, "DialogFrame");
  const rawMemberNames = owner?.sections.flatMap((section) => section.members).map((member) => String(member.name)) ?? [];
  const normalizedMemberNames = rawMemberNames.map((memberName) => normalizeMemberName(memberName));

  assert.ok(owner, "DialogFrame owner が補助ソースとして必要です");

  for (const memberName of rawMemberNames) {
    assert.equal(/^_dummy/iu.test(memberName), false, `DialogFrame.${memberName} は dummy member を含められません`);
    assert.equal(/^_/u.test(memberName), false, `DialogFrame.${memberName} は legacy member を含められません`);
  }

  assert.equal(
    new Set(normalizedMemberNames).size,
    normalizedMemberNames.length,
    "DialogFrame owner に正規化後の重複 member 名があります",
  );
});

test("DialogSheet control collection owner は allow list の property / method だけを保持する", async () => {
  const data = await loadReferenceData();

  for (const config of expectedDialogSheetControlCollectionOwnerConfigs) {
    const { owner, memberNames } = getOwnerMemberNames(data, config.collectionName);
    const normalizedAllowedMemberNames = new Set(["Count", "Item"].map((memberName) => normalizeMemberName(memberName)));

    assert.ok(owner, `${config.collectionName} owner が補助ソースとして必要です`);
    assert.deepEqual(
      [...memberNames].sort(),
      [...normalizedAllowedMemberNames].sort(),
      `${config.collectionName} owner は Count / Item のみを保持する必要があります`,
    );
  }
});

test("DialogSheet control collection owner の Item は literal selector 用の collection type を保持する", async () => {
  const data = await loadReferenceData();

  for (const config of expectedDialogSheetControlCollectionOwnerConfigs) {
    const itemMember = getMember(data, config.collectionName, "Item");
    const itemSignature = getSignature(data, config.collectionName, "Item");

    assert.ok(itemMember, `${config.collectionName}.Item member が必要です`);
    assert.equal(
      itemMember.typeName,
      config.collectionName,
      `${config.collectionName}.Item は literal selector 正規化のため collection owner を保持する必要があります`,
    );
    assert.ok(itemSignature, `${config.collectionName}.Item の署名が必要です`);
    assert.equal(itemSignature.parameters.length, 1, `${config.collectionName}.Item は単一引数の署名である必要があります`);
  }
});

test("DialogSheet control item owner は allow list の property / method だけを保持する", async () => {
  const data = await loadReferenceData();

  for (const config of expectedDialogSheetControlCollectionOwnerConfigs) {
    const { owner, memberNames } = getOwnerMemberNames(data, config.itemName);
    const normalizedAllowedMemberNames = new Set(
      [...config.itemPropertyMemberNames, ...config.itemMethodMemberNames].map((memberName) => normalizeMemberName(memberName)),
    );

    assert.ok(owner, `${config.itemName} owner が補助ソースとして必要です`);
    assert.deepEqual(
      [...memberNames].sort(),
      [...normalizedAllowedMemberNames].sort(),
      `${config.itemName} owner は allow list のみを保持する必要があります`,
    );
  }
});

test("DialogSheet control item owner は代表 property / method を保持する", async () => {
  const data = await loadReferenceData();

  for (const config of expectedDialogSheetControlCollectionOwnerConfigs) {
    const captionMember = getMember(data, config.itemName, "Caption");
    const selectSignature = getSignature(data, config.itemName, "Select");

    assert.ok(captionMember, `${config.itemName}.Caption member が必要です`);
    assert.ok(selectSignature, `${config.itemName}.Select の署名が必要です`);
    assert.equal(selectSignature.ownerName, config.itemName, `${config.itemName}.Select の ownerName が必要です`);
  }
});

test("DialogSheet control owner は dummy / legacy member と正規化重複を含まない", async () => {
  const data = await loadReferenceData();

  for (const ownerName of expectedDialogSheetControlCollectionOwnerConfigs.flatMap((config) => [
    config.collectionName,
    config.itemName,
  ])) {
    const { owner } = getOwnerMemberNames(data, ownerName);
    const rawMemberNames = owner?.sections.flatMap((section) => section.members).map((member) => String(member.name)) ?? [];
    const normalizedMemberNames = rawMemberNames.map((memberName) => normalizeMemberName(memberName));

    assert.ok(owner, `${ownerName} owner が補助ソースとして必要です`);

    for (const memberName of rawMemberNames) {
      assert.equal(/^_dummy/iu.test(memberName), false, `${ownerName}.${memberName} は dummy member を含められません`);
      assert.equal(/^_/u.test(memberName), false, `${ownerName}.${memberName} は legacy member を含められません`);
    }

    assert.equal(
      new Set(normalizedMemberNames).size,
      normalizedMemberNames.length,
      `${ownerName} owner に正規化後の重複 member 名があります`,
    );
  }
});
