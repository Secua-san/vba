import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const outputDir = path.join(rootDir, "resources", "reference");
const outputFile = path.join(outputDir, "mslearn-vba-reference.json");
const apiBaseUrl = "https://learn.microsoft.com/en-us/office/vba/api/";

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
  const response = await fetch(url, {
    headers: {
      Accept: "text/markdown, application/json;q=0.9, text/plain;q=0.8",
    },
  });
  if (!response.ok) {
    throw new Error(`Failed to fetch ${url}: ${response.status} ${response.statusText}`);
  }

  return response.text();
}

async function fetchJson(url) {
  const response = await fetch(url, {
    headers: {
      Accept: "application/json, text/plain;q=0.8",
    },
  });
  if (!response.ok) {
    throw new Error(`Failed to fetch ${url}: ${response.status} ${response.statusText}`);
  }

  return response.json();
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

function parseKeywords(markdown, baseUrl) {
  const table = parseMarkdownTableBlocks(markdown)[0];
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
  const table = parseMarkdownTableBlocks(markdown)[0];
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
  const table = parseMarkdownTableBlocks(markdown)[0];
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

  const output = {
    generatedAt: new Date().toISOString(),
    source: {
      provider: "Microsoft Learn",
      notes: [
        "Excel and Office library reference items are extracted from the official Office VBA API TOC JSON.",
        "Language reference lists are extracted from Microsoft Learn markdown pages for the VBA language reference.",
        "The Excel constants table comes from the Excel.Constants enumeration page on Microsoft Learn.",
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

await main();
