import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { createMcpRequestClient } from "./lib/mcpRequest.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const outputDir = path.join(rootDir, "resources", "reference");
const outputFile = path.join(outputDir, "mslearn-vba-reference.json");
const apiBaseUrl = "https://learn.microsoft.com/en-us/office/vba/api/";
const fetchTimeoutMs = 30_000;
const fetchMinIntervalMs = 250;
const maxFetchRetries = 5;
const signatureMemberAllowList = new Map([
  ["Application", new Set(["Calculate", "CalculateFull", "CalculateFullRebuild", "CalculateUntilAsyncQueriesDone"])],
  [
    "WorksheetFunction",
    new Set([
      "And",
      "Average",
      "Count",
      "CountA",
      "CountBlank",
      "EDate",
      "EoMonth",
      "Find",
      "HLookup",
      "Index",
      "Lookup",
      "Max",
      "Match",
      "Median",
      "Min",
      "Or",
      "Power",
      "Round",
      "Search",
      "Sum",
      "Text",
      "VLookup",
      "Xor",
    ])
  ],
]);
const signatureMetadataOverrides = new Map([
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
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/\\([\\`*_{}\[\]()#+\-.!])/g, "$1")
    .replace(/\s+/g, " ")
    .trim();
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

async function enrichApiMethodSignatures(items) {
  const targets = [];

  for (const item of items) {
    const allowedMembers = signatureMemberAllowList.get(item.name);

    if (!allowedMembers) {
      continue;
    }

    for (const section of item.sections) {
      if (!/^methods$/i.test(section.title)) {
        continue;
      }

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
    (parameter, index) =>
      index > 0 && !parameter.dataType && !parameter.description && parameter.isRequired === undefined,
  );

  if (!hasMissingMetadata) {
    return parameters;
  }

  const templateParameter = parameters.find((parameter) => parameter.dataType || parameter.description);

  if (!templateParameter) {
    return parameters;
  }

  const fallbackDataType = templateParameter.dataType ?? "Variant";

  return parameters.map((parameter, index) => {
    if (index === 0) {
      return parameter;
    }

    if (parameter.dataType || parameter.description || parameter.isRequired !== undefined) {
      return parameter;
    }

    return {
      ...parameter,
      dataType: fallbackDataType,
      description: templateParameter.description,
      isRequired: false,
      label: `${parameter.name} As ${fallbackDataType}`,
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
    parameters[0]?.isRequired !== true ||
    parameters.slice(1).some((parameter) => parameter.isRequired !== true)
  ) {
    return parameters;
  }

  return parameters.map((parameter, index) => ({
    ...parameter,
    isRequired: index === 0,
  }));
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
  return returnType ? `${baseLabel} As ${returnType}` : baseLabel;
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
  const parameters = fillMissingSequentialParameterMetadata(
    syntaxLine,
    applyVariadicTailOptionalRule(
      syntaxLine,
      applySyntaxOptionalParameterRequirements(
        syntaxLine,
        buildSignatureParameterMetadata(signatureParameterNames, parameterTableRows),
      ),
    ),
  );
  const returnType = extractReturnType(returnValueSection);
  const signatureMetadataOverride = signatureMetadataOverrides.get(createSignatureMetadataOverrideKey(ownerName, memberName));
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
      syntaxLine || overriddenParameters.length > 0 || returnType
        ? {
            label: buildSignatureLabel(memberName, signatureParameterNames, returnType),
            ownerName,
            parameters: overriddenParameters,
            returnType,
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

  const output = {
    source: {
      provider: "Microsoft Learn",
      notes: [
        "Excel and Office library reference items are extracted from the official Office VBA API TOC JSON.",
        "Language reference lists are extracted from Microsoft Learn markdown pages for the VBA language reference.",
        "The Excel constants table comes from the Excel.Constants enumeration page on Microsoft Learn.",
        "Method summaries and signature metadata are currently enriched for selected Excel methods.",
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
