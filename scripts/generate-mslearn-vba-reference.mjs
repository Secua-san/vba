import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { createMcpRequestClient } from "./lib/mcpRequest.mjs";
import { signatureMemberAllowList } from "./lib/referenceSignatureConfig.mjs";
import {
  supplementalInteropOwners,
  supplementalOwnerClones,
  supplementalOwnerMembers,
} from "./lib/supplementalReferenceConfig.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const outputDir = path.join(rootDir, "resources", "reference");
const outputFile = path.join(outputDir, "mslearn-vba-reference.json");
const apiBaseUrl = "https://learn.microsoft.com/en-us/office/vba/api/";
const fetchTimeoutMs = 30_000;
const fetchMinIntervalMs = 250;
const maxFetchRetries = 5;
const signatureMetadataOverrides = new Map([
  [
    "workbook.close",
    {
      returnType: "Void",
    },
  ],
  [
    "workbook.exportasfixedformat",
    {
      label:
        "ExportAsFixedFormat(Type, FileName, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish, FixedFormatExtClassPtr)",
      returnType: "Void",
    },
  ],
  [
    "workbook.saveas",
    {
      label:
        "SaveAs(FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode, ConflictResolution, AddToMru, TextCodepage, TextVisualLayout, Local)",
      returnType: "Void",
    },
  ],
  [
    "worksheet.exportasfixedformat",
    {
      label:
        "ExportAsFixedFormat(Type, FileName, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish, FixedFormatExtClassPtr)",
      returnType: "Void",
    },
  ],
  [
    "worksheet.saveas",
    {
      label:
        "SaveAs(FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AddToMru, TextCodepage, TextVisualLayout, Local)",
      returnType: "Void",
    },
  ],
  [
    "worksheetfunction.find",
    {
      parameterDescriptions: new Map([
        ["arg1", "Find_text - the text that you want to find."],
        ["arg2", "Within_text - the text in which you want to search for find_text."],
        ["arg3", "Start_num - the character number in within_text at which you want to start searching."],
      ]),
      summary: "Finds a substring within a text string and returns the starting position.",
    },
  ],
]);
const signatureOwnerNames = new Set(signatureMemberAllowList.keys());
const microsoftLearnClient = createMcpRequestClient({
  baseDelayMs: 2_000,
  maxDelayMs: 60_000,
  maxRetries: maxFetchRetries,
  mcpName: "microsoft-learn",
  minIntervalMs: fetchMinIntervalMs,
  timeoutMs: fetchTimeoutMs,
});

const sourceUrls = {
  apiToc: "https://learn.microsoft.com/en-us/office/vba/api/toc.json",
  excelLanding: "https://learn.microsoft.com/en-us/office/vba/api/overview/excel",
  excelObjectModel: "https://learn.microsoft.com/en-us/office/vba/api/overview/excel/object-model",
  excelConstants: "https://learn.microsoft.com/en-us/office/vba/api/excel.constants",
  languageLanding:
    "https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/visual-basic-language-reference",
  keywords:
    "https://learn.microsoft.com/en-us/office/vba/language/reference/keywords-visual-basic-for-applications",
  constants:
    "https://learn.microsoft.com/en-us/office/vba/language/reference/constants-visual-basic-for-applications",
  functions:
    "https://learn.microsoft.com/en-us/office/vba/language/reference/functions-visual-basic-for-applications",
  operators: "https://learn.microsoft.com/en-us/office/vba/language/reference/operators",
  objects:
    "https://learn.microsoft.com/en-us/office/vba/language/reference/objects-visual-basic-for-applications",
  statements: "https://learn.microsoft.com/en-us/office/vba/language/reference/statements",
  libraryLanding: "https://learn.microsoft.com/en-us/office/vba/api/overview/library-reference",
  libraryReference:
    "https://learn.microsoft.com/en-us/office/vba/api/overview/library-reference/reference-object-library-reference-for-office",
};

const interopReservedSummary = "Reserved for internal use.";

async function fetchText(url) {
  return microsoftLearnClient.request({
    init: {
      headers: {
        Accept: "text/markdown, application/json;q=0.9, text/plain;q=0.8",
      },
    },
    operationName: "fetch-text",
    parseResponse: (response) => response.text(),
    requestKey: `GET text ${url}`,
    url,
  });
}

async function fetchJson(url) {
  return microsoftLearnClient.request({
    init: {
      headers: {
        Accept: "application/json, text/plain;q=0.8",
      },
    },
    operationName: "fetch-json",
    parseResponse: (response) => response.json(),
    requestKey: `GET json ${url}`,
    url,
  });
}

function withMarkdown(url) {
  const value = new URL(url);
  value.searchParams.set("accept", "text/markdown");
  return value.toString();
}

function stripFrontMatter(markdown) {
  const normalized = markdown.replace(/\r\n/g, "\n");
  if (!normalized.startsWith("---\n")) {
    return normalized;
  }

  const closingIndex = normalized.indexOf("\n---\n", 4);
  if (closingIndex === -1) {
    return normalized;
  }

  return normalized.slice(closingIndex + 5);
}

function stripMarkdownText(value) {
  return value
    .replace(/\*\*([^*]+)\*\*/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    .replace(/\[([^\]]+)\]\([^)]+\)/g, "$1")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/\\([\\`*_{}\[\]()#+\-.!])/g, "$1")
    .replace(/([a-z0-9)])\.(?=(?:[A-Z][a-z]|[A-Z]{2,}:))/g, "$1. ")
    .replace(/\s+/g, " ")
    .replace(/\bWorkbooks\.\s+Open\b/giu, "Workbooks.Open")
    .trim();
}

function normalizeReferenceName(value) {
  return String(value ?? "")
    .replace(/\s+/g, "")
    .toLowerCase();
}

function parseInlineLinks(value, baseUrl) {
  const items = [];
  for (const match of value.matchAll(/\[([^\]]+)\]\(([^)]+)\)/g)) {
    items.push({
      name: stripMarkdownText(match[1]),
      learnUrl: new URL(match[2], baseUrl).toString(),
    });
  }

  return items;
}

function parseMarkdownTableBlocks(markdown) {
  const blocks = [];
  const lines = stripFrontMatter(markdown).split("\n");

  for (let index = 0; index < lines.length; index += 1) {
    if (!lines[index].trim().startsWith("|")) {
      continue;
    }

    const block = [];
    while (index < lines.length && lines[index].trim().startsWith("|")) {
      block.push(lines[index].trim());
      index += 1;
    }

    if (block.length < 2) {
      continue;
    }

    const separator = block[1].replace(/[|\-\s:]/g, "");
    if (separator.length !== 0) {
      continue;
    }

    const headers = splitTableRow(block[0]);
    const rows = block.slice(2).map((line) => splitTableRow(line));
    blocks.push({ headers, rows });
  }

  return blocks;
}

function splitTableRow(line) {
  return line
    .replace(/^\|/, "")
    .replace(/\|$/, "")
    .split("|")
    .map((value) => value.trim());
}

function requireFirstTable(markdown, parserName, baseUrl) {
  const table = parseMarkdownTableBlocks(markdown)[0];
  if (!table) {
    throw new Error(`${parserName} could not find a markdown table in ${baseUrl}`);
  }

  return table;
}

function parseSectionedLinks(markdown, baseUrl) {
  const items = [];
  const lines = stripFrontMatter(markdown).split("\n");
  let currentSection = null;

  for (const line of lines) {
    const headingMatch = line.match(/^##\s+(.+)$/);
    if (headingMatch) {
      currentSection = stripMarkdownText(headingMatch[1]);
      continue;
    }

    const itemMatch = line.match(/^\s*(?:-|\d+\.)\s+\[([^\]]+)\]\(([^)]+)\)/);
    if (!itemMatch) {
      continue;
    }

    items.push({
      section: currentSection,
      name: stripMarkdownText(itemMatch[1]),
      learnUrl: new URL(itemMatch[2], baseUrl).toString(),
    });
  }

  return items;
}

function toSectionMap(items) {
  const map = {};
  for (const item of items) {
    const key = item.section || "Overview";
    map[key] ??= [];
    map[key].push({
      name: item.name,
      learnUrl: item.learnUrl,
    });
  }
  return map;
}

function nodeTitle(node) {
  return String(node.toc_title || node.title || "");
}

function walkToc(nodes, path = [], visitor) {
  for (const node of nodes || []) {
    const title = nodeTitle(node);
    const nextPath = title ? [...path, title] : path;
    visitor(node, nextPath);
    walkToc(node.children, nextPath, visitor);
  }
}

function findTocNodeByPath(toc, expectedPath) {
  let result = null;
  walkToc(toc.items || [], [], (node, path) => {
    if (result) {
      return;
    }

    if (path.length !== expectedPath.length) {
      return;
    }

    if (path.every((part, index) => part === expectedPath[index])) {
      result = node;
    }
  });

  return result;
}

function toApiLearnUrl(href) {
  return new URL(href, apiBaseUrl).toString();
}

function normalizeApiItemName(title) {
  return title.replace(/\s+(object|enumeration|collection)$/i, "").trim();
}

function inferApiItemKind(title) {
  if (/enumeration$/i.test(title)) {
    return "enumeration";
  }
  if (/collection$/i.test(title)) {
    return "collection";
  }
  if (/object$/i.test(title)) {
    return "object";
  }
  return "other";
}

function parseApiReferenceSection(node) {
  const items = [];
  const enumerations = [];

  for (const child of node.children || []) {
    const title = nodeTitle(child);
    if (!title || /^overview$/i.test(title)) {
      continue;
    }

    if (/^enumerations$/i.test(title)) {
      for (const enumerationNode of child.children || []) {
        enumerations.push({
          name: normalizeApiItemName(nodeTitle(enumerationNode)),
          title: nodeTitle(enumerationNode),
          kind: "enumeration",
          learnUrl: toApiLearnUrl(enumerationNode.href),
          tocHref: String(enumerationNode.href || ""),
        });
      }
      continue;
    }

    const sections = [];
    for (const sectionNode of child.children || []) {
      const sectionTitle = nodeTitle(sectionNode);
      if (!sectionTitle) {
        continue;
      }

      sections.push({
        title: sectionTitle,
        members: (sectionNode.children || []).map((memberNode) => ({
          name: nodeTitle(memberNode),
          learnUrl: toApiLearnUrl(memberNode.href),
          tocHref: String(memberNode.href || ""),
        })),
      });
    }

    items.push({
      name: normalizeApiItemName(title),
      title,
      kind: inferApiItemKind(title),
      learnUrl: toApiLearnUrl(child.href),
      tocHref: String(child.href || ""),
      sections,
    });
  }

  return {
    items,
    enumerations,
  };
}

function findApiReferenceItem(items, itemName) {
  return items.find((item) => normalizeReferenceName(item.name) === normalizeReferenceName(itemName));
}

function cloneApiReferenceItem(sourceItem, cloneConfig) {
  const memberTypeOverrides = cloneConfig.memberTypeOverrides ?? new Map();
  const unresolvedMemberTypeOverrides = new Map(
    [...memberTypeOverrides].map(([memberName, typeName]) => [normalizeReferenceName(memberName), typeName]),
  );

  const clonedItem = {
    kind: cloneConfig.kind ?? sourceItem.kind,
    learnUrl: cloneConfig.learnUrl ?? sourceItem.learnUrl,
    name: cloneConfig.name,
    sections: sourceItem.sections.map((section) => ({
      ...section,
      members: section.members.map((member) => {
        const normalizedMemberName = normalizeReferenceName(member.name);
        const overriddenTypeName = unresolvedMemberTypeOverrides.get(normalizedMemberName);

        if (overriddenTypeName) {
          unresolvedMemberTypeOverrides.delete(normalizedMemberName);
        }

        return overriddenTypeName
          ? {
              ...member,
              typeName: overriddenTypeName,
            }
          : { ...member };
      }),
    })),
    title: cloneConfig.title ?? sourceItem.title,
    tocHref: cloneConfig.tocHref ?? "",
  };

  if (unresolvedMemberTypeOverrides.size > 0) {
    const [memberName] = unresolvedMemberTypeOverrides.keys();
    throw new Error(`Supplemental owner clone '${cloneConfig.name}' could not resolve member type override '${memberName}'.`);
  }

  return clonedItem;
}

function applySupplementalOwnerMembers(items) {
  for (const ownerConfig of supplementalOwnerMembers) {
    const targetOwner = findApiReferenceItem(items, ownerConfig.ownerName);

    if (!targetOwner) {
      throw new Error(`Supplemental owner member target '${ownerConfig.ownerName}' was not found.`);
    }

    const targetSection = targetOwner.sections.find(
      (section) => normalizeReferenceName(section.title) === normalizeReferenceName(ownerConfig.sectionName),
    );

    if (!targetSection) {
      throw new Error(
        `Supplemental owner member target '${ownerConfig.ownerName}.${ownerConfig.sectionName}' was not found.`,
      );
    }

    const existingMemberNames = new Set(targetSection.members.map((member) => normalizeReferenceName(member.name)));

    for (const member of ownerConfig.members) {
      const normalizedMemberName = normalizeReferenceName(member.name);

      if (existingMemberNames.has(normalizedMemberName)) {
        throw new Error(
          `Supplemental owner member '${ownerConfig.ownerName}.${member.name}' already exists in the target section.`,
        );
      }

      existingMemberNames.add(normalizedMemberName);
      targetSection.members.push({ ...member });
    }
  }
}

function normalizeInteropMemberDisplayName(value) {
  return stripMarkdownText(value).replace(/\(.*$/u, "").trim();
}

function parseSectionLinkTable(markdown, headingTitle, baseUrl) {
  const sectionMarkdown = extractMarkdownSection(markdown, headingTitle);
  const table = parseMarkdownTableBlocks(sectionMarkdown)[0];

  if (!table) {
    return [];
  }

  return table.rows.flatMap((cells) => {
    const [linkCell = "", noteCell = ""] = cells;
    const links = parseInlineLinks(linkCell, baseUrl);
    const note = stripMarkdownText(noteCell) || undefined;

    return links.map((link) => ({
      ...link,
      name: normalizeInteropMemberDisplayName(link.name),
      note,
    }));
  });
}

function extractCodeFence(markdown, language) {
  const match = markdown.match(new RegExp("```" + language + "\\r?\\n([\\s\\S]*?)\\r?\\n```", "iu"));
  return match?.[1]?.trim();
}

function extractInteropVbSignatureLine(definitionSection) {
  const vbBlock = extractCodeFence(definitionSection, "vb");

  if (!vbBlock) {
    return undefined;
  }

  return vbBlock
    .split(/\r?\n/u)
    .map((line) => line.trim())
    .find((line) => line.length > 0);
}

function splitSignatureParameterList(parameterList) {
  if (!parameterList || parameterList.trim().length === 0) {
    return [];
  }

  return parameterList.split(",").map((value) => value.trim()).filter((value) => value.length > 0);
}

function parseInteropSignatureLine(vbSignatureLine) {
  if (!vbSignatureLine) {
    return undefined;
  }

  const normalizedLine = vbSignatureLine.replace(/\s+/g, " ").trim();
  const signatureMatch =
    normalizedLine.match(
      /^Public\s+(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:\((.*)\))?(?:\s+As\s+([A-Za-z_][A-Za-z0-9_.]*))?$/iu,
    ) ?? normalizedLine.match(/^(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:\((.*)\))?(?:\s+As\s+([A-Za-z_][A-Za-z0-9_.]*))?$/iu);

  if (!signatureMatch?.[1] || !signatureMatch[2]) {
    return undefined;
  }

  const procedureKind = signatureMatch[1];
  const memberName = signatureMatch[2];
  const parameterList = signatureMatch[3] ?? "";
  const returnType = procedureKind.toLowerCase() === "function" ? signatureMatch[4] ?? "Variant" : "Void";
  const parameters = splitSignatureParameterList(parameterList).map((rawParameter) => parseInteropSignatureParameter(rawParameter));

  return {
    memberName,
    parameters,
    returnType,
  };
}

function parseInteropSignatureParameter(rawParameter) {
  const cleanedParameter = rawParameter.replace(/\s+/g, " ").trim();
  const isRequired = !/^Optional\b/iu.test(cleanedParameter);
  const withoutOptional = cleanedParameter.replace(/^Optional\s+/iu, "");
  const withoutDirection = withoutOptional.replace(/^(ByRef|ByVal|ParamArray)\s+/iu, "");
  const parameterMatch = withoutDirection.match(/^([A-Za-z_][A-Za-z0-9_]*)\s+As\s+(.+)$/iu);
  const name = parameterMatch?.[1] ?? withoutDirection;
  const dataType = parameterMatch?.[2]?.trim();

  return {
    dataType,
    description: undefined,
    isRequired,
    label: dataType ? `${name} As ${dataType}` : name,
    name,
  };
}

function parseInteropMethodReference(markdown, ownerName, memberName) {
  const definitionSection = extractMarkdownSection(markdown, "Definition");
  const summary = extractSummaryParagraph(markdown);
  const parsedSignature = parseInteropSignatureLine(extractInteropVbSignatureLine(definitionSection));

  if (!parsedSignature || normalizeReferenceName(parsedSignature.memberName) !== normalizeReferenceName(memberName)) {
    return {
      signature: undefined,
      summary: summary && summary !== interopReservedSummary ? summary : undefined,
    };
  }

  return {
    signature: {
      label: buildSignatureLabel(memberName, parsedSignature.parameters.map((parameter) => parameter.name), parsedSignature.returnType),
      ownerName,
      parameters: parsedSignature.parameters,
      returnType: parsedSignature.returnType,
    },
    summary: summary && summary !== interopReservedSummary ? summary : undefined,
  };
}

function parseInteropPropertySignatureLine(vbSignatureLine) {
  if (!vbSignatureLine) {
    return undefined;
  }

  const normalizedLine = vbSignatureLine.replace(/\s+/g, " ").trim();
  const signatureMatch =
    normalizedLine.match(
      /^Public\s+(?:(?:ReadOnly|WriteOnly)\s+)?Property\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:\((.*)\))?(?:\s+As\s+([A-Za-z_][A-Za-z0-9_.]*))?$/iu,
    ) ??
    normalizedLine.match(
      /^(?:(?:ReadOnly|WriteOnly)\s+)?Property\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:\((.*)\))?(?:\s+As\s+([A-Za-z_][A-Za-z0-9_.]*))?$/iu,
    );

  if (!signatureMatch?.[1]) {
    return undefined;
  }

  const parameterList = signatureMatch[2] ?? "";

  return {
    memberName: signatureMatch[1],
    parameters: splitSignatureParameterList(parameterList).map((rawParameter) => parseInteropSignatureParameter(rawParameter)),
    returnType: signatureMatch[3] ?? "Variant",
  };
}

function parseInteropPropertyReference(markdown, ownerName, memberName) {
  const definitionSection = extractMarkdownSection(markdown, "Definition");
  const summary = extractSummaryParagraph(markdown);
  const parsedSignature = parseInteropPropertySignatureLine(extractInteropVbSignatureLine(definitionSection));

  if (!parsedSignature || normalizeReferenceName(parsedSignature.memberName) !== normalizeReferenceName(memberName)) {
    return undefined;
  }

  if (!parsedSignature.returnType) {
    return undefined;
  }

  return {
    signature:
      parsedSignature.parameters.length > 0
        ? {
            label: buildSignatureLabel(
              memberName,
              parsedSignature.parameters.map((parameter) => parameter.name),
              parsedSignature.returnType,
            ),
            ownerName,
            parameters: parsedSignature.parameters,
            returnType: parsedSignature.returnType,
          }
        : undefined,
    summary: summary && summary !== interopReservedSummary ? summary : undefined,
    typeName: parsedSignature.returnType,
  };
}

function parseInteropMemberReference(markdown, ownerName, memberName, sectionName) {
  return normalizeReferenceName(sectionName) === "properties"
    ? parseInteropPropertyReference(markdown, ownerName, memberName)
    : parseInteropMethodReference(markdown, ownerName, memberName);
}

function resolveSupplementalInteropOwnerSections(ownerConfig) {
  if (Array.isArray(ownerConfig.sections) && ownerConfig.sections.length > 0) {
    return ownerConfig.sections;
  }

  if (ownerConfig.sectionName && ownerConfig.memberAllowList) {
    return [
      {
        memberAllowList: ownerConfig.memberAllowList,
        sectionName: ownerConfig.sectionName,
      },
    ];
  }

  throw new Error(`Supplemental interop owner '${ownerConfig.name}' must declare at least one section.`);
}

async function buildSupplementalInteropOwner(ownerConfig) {
  const ownerMarkdown = await fetchText(withMarkdown(ownerConfig.learnUrl));
  const seenMemberNames = new Set();
  const sections = [];

  for (const sectionConfig of resolveSupplementalInteropOwnerSections(ownerConfig)) {
    const allowedMemberNames = new Set(
      [...sectionConfig.memberAllowList].map((memberName) => normalizeReferenceName(memberName)),
    );
    const interfaceMembers = parseSectionLinkTable(ownerMarkdown, sectionConfig.sectionName, ownerConfig.learnUrl)
      .filter((member) => !normalizeReferenceName(member.name).startsWith("_"))
      .filter((member) => !normalizeReferenceName(member.name).startsWith("dummy"))
      .filter((member) => allowedMemberNames.has(normalizeReferenceName(member.name)));
    const members = [];

    for (const member of interfaceMembers) {
      const normalizedMemberName = normalizeReferenceName(member.name);

      if (seenMemberNames.has(normalizedMemberName)) {
        throw new Error(`Supplemental interop owner '${ownerConfig.name}' contains duplicate member '${member.name}'.`);
      }

      seenMemberNames.add(normalizedMemberName);
    }

    for (const allowedMemberName of allowedMemberNames) {
      if (!interfaceMembers.some((member) => normalizeReferenceName(member.name) === allowedMemberName)) {
        throw new Error(
          `Supplemental interop owner '${ownerConfig.name}' is missing member '${allowedMemberName}' in section '${sectionConfig.sectionName}'.`,
        );
      }
    }

    for (const member of interfaceMembers) {
      const memberMarkdown = await fetchText(withMarkdown(member.learnUrl));
      const memberMetadata = parseInteropMemberReference(
        memberMarkdown,
        ownerConfig.name,
        member.name,
        sectionConfig.sectionName,
      );

      if (normalizeReferenceName(sectionConfig.sectionName) === "methods" && !memberMetadata.signature) {
        throw new Error(
          `Supplemental interop owner '${ownerConfig.name}' could not extract signature metadata for member '${member.name}'.`,
        );
      }

      if (normalizeReferenceName(sectionConfig.sectionName) === "properties" && !memberMetadata?.typeName) {
        throw new Error(
          `Supplemental interop owner '${ownerConfig.name}' could not extract property type metadata for member '${member.name}'.`,
        );
      }

      members.push({
        learnUrl: member.learnUrl,
        name: member.name,
        signature: memberMetadata?.signature,
        summary: memberMetadata?.summary,
        typeName: memberMetadata?.typeName,
      });
    }

    sections.push({
      members,
      title: sectionConfig.sectionName,
    });
  }

  return {
    kind: ownerConfig.kind,
    learnUrl: ownerConfig.learnUrl,
    name: ownerConfig.name,
    sections,
    title: ownerConfig.title,
    tocHref: "",
  };
}

async function buildSupplementalExcelItems(items) {
  const supplementalItems = [];

  for (const cloneConfig of supplementalOwnerClones) {
    const sourceItem = findApiReferenceItem(items, cloneConfig.sourceOwnerName);

    if (!sourceItem) {
      throw new Error(`Supplemental owner clone source '${cloneConfig.sourceOwnerName}' was not found.`);
    }

    supplementalItems.push(cloneApiReferenceItem(sourceItem, cloneConfig));
  }

  for (const ownerConfig of supplementalInteropOwners) {
    supplementalItems.push(await buildSupplementalInteropOwner(ownerConfig));
  }

  return supplementalItems;
}

async function enrichApiMethodSignatures(items) {
  const targets = [];

  for (const item of items) {
    const allowedMembers = signatureMemberAllowList.get(item.name);

    if (!allowedMembers) {
      continue;
    }

    for (const section of item.sections) {
      for (const member of section.members) {
        if (!member.learnUrl || !allowedMembers.has(member.name)) {
          continue;
        }

        targets.push({
          member,
          ownerName: item.name,
        });
      }
    }
  }

  for (const target of targets) {
    const markdown = await fetchText(withMarkdown(target.member.learnUrl));
    const signatureMetadata = parseApiMethodReference(markdown, target.ownerName, target.member.name);

    if (signatureMetadata.summary) {
      target.member.summary = signatureMetadata.summary;
    }

    if (signatureMetadata.signature) {
      target.member.signature = signatureMetadata.signature;
    }
  }

  return {
    attempted: targets.length,
    resolved: targets.filter((target) => Boolean(target.member.signature)).length,
  };
}

function summarizeApiItems(items, enumerations) {
  const counts = {
    total: items.length + enumerations.length,
    objects: 0,
    collections: 0,
    enumerations: enumerations.length,
    others: 0,
    members: 0,
  };

  for (const item of items) {
    if (item.kind === "object") {
      counts.objects += 1;
    } else if (item.kind === "collection") {
      counts.collections += 1;
    } else if (item.kind === "enumeration") {
      counts.enumerations += 1;
    } else {
      counts.others += 1;
    }

    for (const section of item.sections) {
      counts.members += section.members.length;
    }
  }

  return counts;
}

function extractSummaryParagraph(markdown) {
  const lines = stripFrontMatter(markdown).split("\n");
  let titleSeen = false;
  const paragraph = [];

  for (const rawLine of lines) {
    const line = rawLine.trim();

    if (!titleSeen) {
      if (/^#\s+/u.test(line)) {
        titleSeen = true;
      }
      continue;
    }

    if (line.length === 0) {
      if (paragraph.length > 0) {
        break;
      }
      continue;
    }

    if (/^##\s+/u.test(line)) {
      break;
    }

    paragraph.push(line);
  }

  return paragraph.length > 0 ? stripMarkdownText(paragraph.join(" ")) : undefined;
}

function extractMarkdownSection(markdown, headingTitle) {
  const lines = stripFrontMatter(markdown).split("\n");
  const normalizedHeading = headingTitle.trim().toLowerCase();
  const sectionLines = [];
  let collecting = false;

  for (const rawLine of lines) {
    const headingMatch = rawLine.match(/^##\s+(.+)$/u);

    if (headingMatch) {
      const currentHeading = stripMarkdownText(headingMatch[1]).toLowerCase();

      if (collecting) {
        break;
      }

      if (currentHeading === normalizedHeading) {
        collecting = true;
      }

      continue;
    }

    if (collecting) {
      sectionLines.push(rawLine);
    }
  }

  return sectionLines.join("\n").trim();
}

function extractSyntaxLine(sectionMarkdown) {
  if (!sectionMarkdown) {
    return undefined;
  }

  const lines = sectionMarkdown.split("\n");

  for (const rawLine of lines) {
    const line = rawLine.trim();

    if (line.length === 0) {
      continue;
    }

    return line;
  }

  return undefined;
}

function extractReturnType(sectionMarkdown) {
  if (!sectionMarkdown) {
    return undefined;
  }

  for (const rawLine of sectionMarkdown.split("\n")) {
    const line = stripMarkdownText(rawLine);

    if (line.length > 0) {
      return line;
    }
  }

  return undefined;
}

function inferReturnTypeFromSummary(summary) {
  if (!summary) {
    return undefined;
  }

  const readOnlyMatch = summary.match(/\bRead-only\s+([A-Z][A-Za-z0-9_]*)\b/u);

  if (readOnlyMatch?.[1]) {
    return readOnlyMatch[1];
  }

  const returnsValueMatch = summary.match(/\bReturns\s+(?:an?\s+)?([A-Z][A-Za-z0-9_]*)\s+value\b/u);

  if (returnsValueMatch?.[1]) {
    return returnsValueMatch[1];
  }

  const returnsTypeMatch = summary.match(/\bReturns\s+(?:an?\s+)?([A-Z][A-Za-z0-9_]*)\b/u);
  return returnsTypeMatch?.[1];
}

function extractSyntaxParameterNames(syntaxLine) {
  if (!syntaxLine) {
    return [];
  }

  const openParenIndex = syntaxLine.indexOf("(");
  const closeParenIndex = syntaxLine.lastIndexOf(")");

  if (openParenIndex === -1 || closeParenIndex <= openParenIndex) {
    return [];
  }

  return syntaxLine
    .slice(openParenIndex + 1, closeParenIndex)
    .split(",")
    .map((value) => normalizeSignatureParameterName(value))
    .filter((value) => value.length > 0);
}

function normalizeSignatureParameterName(value) {
  return stripMarkdownText(value)
    .replace(/^\*+|\*+$/g, "")
    .replace(/^\[+|\]+$/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function parseParameterTableRows(sectionMarkdown) {
  const table = parseMarkdownTableBlocks(sectionMarkdown)[0];

  if (!table) {
    return [];
  }

  return table.rows.map((cells) => {
    const [nameCell = "", requiredCell = "", dataTypeCell = "", descriptionCell = ""] = cells;
    return {
      dataType: stripMarkdownText(dataTypeCell) || undefined,
      description: stripMarkdownText(descriptionCell) || undefined,
      isRequired: /^required$/iu.test(stripMarkdownText(requiredCell)),
      name: normalizeSignatureParameterName(nameCell),
    };
  });
}

function expandSignatureParameterNames(parameterName, syntaxParameterNames) {
  if (!parameterName) {
    return [];
  }

  if (syntaxParameterNames.includes(parameterName)) {
    return [parameterName];
  }

  const compactParameterName = parameterName.replace(/\s+/g, "");
  const rangeMatch =
    compactParameterName.match(/^([A-Za-z_]+)(\d+)(?:,?(?:\.{3}|…|-),?)(?:([A-Za-z_]+))?(\d+)$/u) ??
    parameterName.match(/^([A-Za-z_]+)(\d+)\s*-\s*(?:([A-Za-z_]+))?(\d+)$/u);

  if (!rangeMatch) {
    return [parameterName];
  }

  const [, startPrefix, startValueText, endPrefix = startPrefix, endValueText] = rangeMatch;

  if (startPrefix.toLowerCase() !== endPrefix.toLowerCase()) {
    return [parameterName];
  }

  const startValue = Number(startValueText);
  const endValue = Number(endValueText);

  if (!Number.isInteger(startValue) || !Number.isInteger(endValue) || startValue > endValue) {
    return [parameterName];
  }

  return Array.from({ length: endValue - startValue + 1 }, (_, index) => `${startPrefix}${startValue + index}`);
}

function getTrailingNumericSuffix(value) {
  const match = value.match(/(\d+)$/u);
  return match?.[1];
}

function hasVariadicSyntaxMarker(syntaxLine) {
  return /(?:\.{3}|…)/u.test(syntaxLine ?? "");
}

function fillMissingSequentialParameterMetadata(syntaxLine, parameters) {
  if (!hasSequentialNumericSuffixParameters(parameters) || parameters.length < 10) {
    return parameters;
  }

  const hasMissingMetadata = parameters.some(
    (parameter) =>
      !parameter.dataType || !parameter.description || parameter.isRequired === undefined || parameter.label === parameter.name,
  );

  if (!hasMissingMetadata) {
    return parameters;
  }

  const templateParameter = parameters.find((parameter) => parameter.dataType || parameter.description);

  if (!templateParameter) {
    return parameters;
  }

  const fallbackDataType = templateParameter.dataType ?? "Variant";
  const optionalParameterNames = extractOptionalSyntaxParameterNames(syntaxLine);

  return parameters.map((parameter, index) => {
    const dataType = parameter.dataType ?? fallbackDataType;
    const description = parameter.description ?? templateParameter.description;
    const isRequired =
      parameter.isRequired ??
      (optionalParameterNames.has(parameter.name)
        ? false
        : index === 0
          ? true
          : templateParameter.isRequired ?? false);
    const label = dataType ? `${parameter.name} As ${dataType}` : parameter.label;

    if (
      parameter.dataType === dataType &&
      parameter.description === description &&
      parameter.isRequired === isRequired &&
      parameter.label === label
    ) {
      return parameter;
    }

    return {
      ...parameter,
      dataType,
      description,
      isRequired,
      label,
    };
  });
}

function buildSignatureParameterMetadata(syntaxParameterNames, tableRows) {
  const rowEntries = tableRows.flatMap((row) =>
    expandSignatureParameterNames(row.name, syntaxParameterNames).map((parameterName) => [parameterName, row]),
  );
  const metadataByName = new Map(rowEntries);
  const metadataByNumericSuffix = new Map();

  for (const [parameterName, row] of rowEntries) {
    const numericSuffix = getTrailingNumericSuffix(parameterName);

    if (!numericSuffix) {
      continue;
    }

    const existing = metadataByNumericSuffix.get(numericSuffix);

    if (!existing) {
      metadataByNumericSuffix.set(numericSuffix, row);
      continue;
    }

    if (existing !== row) {
      metadataByNumericSuffix.set(numericSuffix, null);
    }
  }

  return syntaxParameterNames.map((parameterName) => {
    const numericSuffix = getTrailingNumericSuffix(parameterName);
    const metadataFromSuffix = numericSuffix ? metadataByNumericSuffix.get(numericSuffix) : undefined;
    const metadata = metadataByName.get(parameterName) ?? (metadataFromSuffix && metadataFromSuffix !== null ? metadataFromSuffix : undefined);
    const dataType = metadata?.dataType;
    const label = dataType ? `${parameterName} As ${dataType}` : parameterName;
    return {
      dataType,
      description: metadata?.description,
      isRequired: metadata?.isRequired,
      label,
      name: parameterName,
    };
  });
}

function expandTableParameterNames(tableRows, syntaxParameterNames) {
  const names = [];

  for (const row of tableRows) {
    for (const parameterName of expandSignatureParameterNames(row.name, syntaxParameterNames)) {
      if (parameterName.length === 0 || parameterName === "..." || parameterName === "…" || names.includes(parameterName)) {
        continue;
      }

      names.push(parameterName);
    }
  }

  return names;
}

function resolveSignatureParameterNames(syntaxParameterNames, tableRows) {
  const syntaxNames = syntaxParameterNames.filter(
    (parameterName) => parameterName.length > 0 && parameterName !== "..." && parameterName !== "…",
  );
  const hasVariadicMarker = syntaxParameterNames.includes("...") || syntaxParameterNames.includes("…");

  if (!hasVariadicMarker) {
    return syntaxNames;
  }

  const expandedTableNames = expandTableParameterNames(tableRows, syntaxNames);
  return expandedTableNames.length > syntaxNames.length ? expandedTableNames : syntaxNames;
}

function extractOptionalSyntaxParameterNames(syntaxLine) {
  if (!syntaxLine) {
    return new Set();
  }

  const openParenIndex = syntaxLine.indexOf("(");
  const closeParenIndex = syntaxLine.lastIndexOf(")");

  if (openParenIndex === -1 || closeParenIndex <= openParenIndex) {
    return new Set();
  }

  return new Set(
    syntaxLine
      .slice(openParenIndex + 1, closeParenIndex)
      .split(",")
      .map((rawValue) => rawValue.trim())
      .filter((rawValue) => rawValue.startsWith("[") && rawValue.endsWith("]"))
      .map((rawValue) => normalizeSignatureParameterName(rawValue))
      .filter((value) => value.length > 0),
  );
}

function applySyntaxOptionalParameterRequirements(syntaxLine, parameters) {
  const optionalParameterNames = extractOptionalSyntaxParameterNames(syntaxLine);

  if (optionalParameterNames.size === 0) {
    return parameters;
  }

  return parameters.map((parameter) =>
    optionalParameterNames.has(parameter.name)
      ? {
          ...parameter,
          isRequired: false,
        }
      : parameter,
  );
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

function hasVariadicCountDescription(parameters) {
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

function applyVariadicTailOptionalRule(syntaxLine, parameters) {
  if (
    (!hasVariadicSyntaxMarker(syntaxLine) && !hasVariadicCountDescription(parameters)) ||
    !hasSequentialNumericSuffixParameters(parameters) ||
    parameters[0]?.isRequired !== true
  ) {
    return parameters;
  }

  return parameters.map((parameter, index) =>
    // Variadic な sequential parameters は先頭だけ必須にそろえ、既に期待値どおりなら再生成しない。
    parameter.isRequired === (index === 0)
      ? parameter
      : {
          ...parameter,
          isRequired: index === 0,
        },
  );
}

function summarizeSignatureLabelParameters(parameterNames) {
  if (parameterNames.length <= 6) {
    return parameterNames;
  }

  return [...parameterNames.slice(0, 3), "...", parameterNames[parameterNames.length - 1]];
}

function buildSignatureLabel(memberName, syntaxParameterNames, returnType) {
  const parameterList = summarizeSignatureLabelParameters(syntaxParameterNames).join(", ");
  const baseLabel = `${memberName}(${parameterList})`;
  return returnType && returnType !== "Void" ? `${baseLabel} As ${returnType}` : baseLabel;
}

function createSignatureMetadataOverrideKey(ownerName, memberName) {
  return `${ownerName}.${memberName}`.replace(/\s+/g, "").toLowerCase();
}

function parseApiMethodReference(markdown, ownerName, memberName) {
  const summary = extractSummaryParagraph(markdown);
  const syntaxSection = extractMarkdownSection(markdown, "Syntax");
  const parametersSection = extractMarkdownSection(markdown, "Parameters");
  const returnValueSection = extractMarkdownSection(markdown, "Return value");
  const syntaxLine = extractSyntaxLine(syntaxSection);
  const syntaxParameterNames = extractSyntaxParameterNames(syntaxLine);
  const parameterTableRows = parseParameterTableRows(parametersSection);
  const signatureParameterNames = resolveSignatureParameterNames(syntaxParameterNames, parameterTableRows);
  const parameters = applyVariadicTailOptionalRule(
    syntaxLine,
    fillMissingSequentialParameterMetadata(
      syntaxLine,
      applySyntaxOptionalParameterRequirements(
        syntaxLine,
        buildSignatureParameterMetadata(signatureParameterNames, parameterTableRows),
      ),
    ),
  );
  const returnType = extractReturnType(returnValueSection) ?? inferReturnTypeFromSummary(summary);
  const signatureMetadataOverride = signatureMetadataOverrides.get(createSignatureMetadataOverrideKey(ownerName, memberName));
  const resolvedReturnType = signatureMetadataOverride?.returnType ?? returnType;
  const overriddenParameters = signatureMetadataOverride
    ? parameters.map((parameter) => ({
        ...parameter,
        description:
          signatureMetadataOverride.parameterDescriptions?.get(
            normalizeSignatureParameterName(parameter.name).toLowerCase(),
          ) ?? parameter.description,
      }))
    : parameters;
  const resolvedSummary = signatureMetadataOverride?.summary ?? summary;

  return {
    signature:
      syntaxLine || overriddenParameters.length > 0 || resolvedReturnType
        ? {
            label:
              signatureMetadataOverride?.label ??
              buildSignatureLabel(memberName, signatureParameterNames, resolvedReturnType),
            ownerName,
            parameters: overriddenParameters,
            returnType: resolvedReturnType,
          }
        : undefined,
    summary: resolvedSummary,
  };
}

function parseKeywords(markdown, baseUrl) {
  const table = requireFirstTable(markdown, "parseKeywords", baseUrl);
  return table.rows.map((cells) => {
    const [keywordCell = "", contextsCell = ""] = cells;
    const contexts = parseInlineLinks(contextsCell, baseUrl);
    const note = contexts.length === 0 ? stripMarkdownText(contextsCell) : null;

    return {
      name: stripMarkdownText(keywordCell),
      contexts,
      note,
    };
  });
}

function parseOperatorSummary(markdown, baseUrl) {
  const table = requireFirstTable(markdown, "parseOperatorSummary", baseUrl);
  return table.rows.map((cells) => {
    const [groupCell = "", descriptionCell = "", operatorsCell = ""] = cells;
    const groupLinks = parseInlineLinks(groupCell, baseUrl);
    return {
      group: stripMarkdownText(groupCell),
      description: stripMarkdownText(descriptionCell),
      learnUrl: groupLinks[0]?.learnUrl ?? null,
      operators: parseInlineLinks(operatorsCell, baseUrl),
    };
  });
}

function parseExcelConstants(markdown, baseUrl) {
  const table = requireFirstTable(markdown, "parseExcelConstants", baseUrl);
  return table.rows.map((cells) => {
    const [nameCell = "", valueCell = "", descriptionCell = ""] = cells;
    return {
      name: stripMarkdownText(nameCell),
      value: stripMarkdownText(valueCell),
      description: stripMarkdownText(descriptionCell),
      learnUrl: baseUrl,
    };
  });
}

function parseLandingSections(markdown, baseUrl) {
  return parseSectionedLinks(markdown, baseUrl).map((item) => ({
    name: item.name,
    learnUrl: item.learnUrl,
  }));
}

async function main() {
  const [
    toc,
    languageLandingMarkdown,
    keywordsMarkdown,
    constantsMarkdown,
    functionsMarkdown,
    operatorsMarkdown,
    objectsMarkdown,
    statementsMarkdown,
    excelConstantsMarkdown,
  ] = await Promise.all([
    fetchJson(sourceUrls.apiToc),
    fetchText(withMarkdown(sourceUrls.languageLanding)),
    fetchText(withMarkdown(sourceUrls.keywords)),
    fetchText(withMarkdown(sourceUrls.constants)),
    fetchText(withMarkdown(sourceUrls.functions)),
    fetchText(withMarkdown(sourceUrls.operators)),
    fetchText(withMarkdown(sourceUrls.objects)),
    fetchText(withMarkdown(sourceUrls.statements)),
    fetchText(withMarkdown(sourceUrls.excelConstants)),
  ]);

  const excelObjectModelNode = findTocNodeByPath(toc, ["Office VBA Reference", "Excel", "Object model"]);
  const libraryReferenceNode = findTocNodeByPath(toc, [
    "Office VBA Reference",
    "Library reference",
    "Reference",
  ]);

  if (!excelObjectModelNode || !libraryReferenceNode) {
    throw new Error("Could not locate Excel or library reference sections in Microsoft Learn TOC.");
  }

  const excelReference = parseApiReferenceSection(excelObjectModelNode);
  const libraryReference = parseApiReferenceSection(libraryReferenceNode);
  const languageFunctions = toSectionMap(parseSectionedLinks(functionsMarkdown, sourceUrls.functions));
  const signatureStats = await enrichApiMethodSignatures(excelReference.items);
  const supplementalExcelItems = await buildSupplementalExcelItems(excelReference.items);
  const supplementalSignatureCount = supplementalExcelItems.reduce(
    (count, item) => count + item.sections.flatMap((section) => section.members).filter((member) => Boolean(member.signature)).length,
    0,
  );
  excelReference.items.push(...supplementalExcelItems);
  applySupplementalOwnerMembers(excelReference.items);

  const output = {
    source: {
      provider: "Microsoft Learn",
      notes: [
        "Excel and Office library reference items are extracted from the official Office VBA API TOC JSON.",
        "Language reference lists are extracted from Microsoft Learn markdown pages for the VBA language reference.",
        "The Excel constants table comes from the Excel.Constants enumeration page on Microsoft Learn.",
        "Member summaries and signature metadata are currently enriched for selected Excel members.",
        "DialogSheet and DialogFrame supplemental members are added from Microsoft Learn interop pages as constrained supplemental sources.",
        "Application.DialogSheets and Workbook.DialogSheets are normalized to the supplemental DialogSheets collection owner for built-in chain resolution.",
      ],
      urls: sourceUrls,
    },
    excel: {
      landingUrl: sourceUrls.excelLanding,
      objectModelUrl: sourceUrls.excelObjectModel,
      constantsEnumerationUrl: sourceUrls.excelConstants,
      objectModel: {
        counts: summarizeApiItems(excelReference.items, excelReference.enumerations),
        items: excelReference.items,
        enumerations: excelReference.enumerations,
      },
      constantsEnumeration: parseExcelConstants(excelConstantsMarkdown, sourceUrls.excelConstants),
    },
    languageReference: {
      landingUrl: sourceUrls.languageLanding,
      landingSections: parseLandingSections(languageLandingMarkdown, sourceUrls.languageLanding),
      keywords: parseKeywords(keywordsMarkdown, sourceUrls.keywords),
      constantCategories: parseSectionedLinks(constantsMarkdown, sourceUrls.constants).map((item) => ({
        name: item.name,
        section: item.section,
        learnUrl: item.learnUrl,
      })),
      functions: languageFunctions,
      operators: parseOperatorSummary(operatorsMarkdown, sourceUrls.operators),
      objects: parseSectionedLinks(objectsMarkdown, sourceUrls.objects).map((item) => ({
        name: item.name,
        section: item.section,
        learnUrl: item.learnUrl,
      })),
      statements: parseSectionedLinks(statementsMarkdown, sourceUrls.statements).map((item) => ({
        name: item.name,
        section: item.section,
        learnUrl: item.learnUrl,
      })),
    },
    libraryReference: {
      landingUrl: sourceUrls.libraryLanding,
      referenceUrl: sourceUrls.libraryReference,
      reference: {
        counts: summarizeApiItems(libraryReference.items, libraryReference.enumerations),
        items: libraryReference.items,
        enumerations: libraryReference.enumerations,
      },
    },
  };

  await mkdir(outputDir, { recursive: true });
  await writeFile(outputFile, `${JSON.stringify(output, null, 2)}\n`, "utf8");

  console.log(
    JSON.stringify(
      {
        outputFile,
        excel: output.excel.objectModel.counts,
        excelSignatureOwners: [...signatureOwnerNames],
        excelSignatures: signatureStats,
        excelSupplementalSignatures: supplementalSignatureCount,
        excelConstants: output.excel.constantsEnumeration.length,
        language: {
          landingSections: output.languageReference.landingSections.length,
          keywords: output.languageReference.keywords.length,
          constantCategories: output.languageReference.constantCategories.length,
          functionSections: Object.keys(output.languageReference.functions).length,
          objects: output.languageReference.objects.length,
          statements: output.languageReference.statements.length,
          operators: output.languageReference.operators.length,
        },
        library: output.libraryReference.reference.counts,
      },
      null,
      2,
    ),
  );
}

main().catch((error) => {
  console.error("Failed to generate Microsoft Learn VBA reference data.", error);
  process.exit(1);
});
