import path from "node:path";

import JSZip from "jszip";
import { parseStringPromise } from "xml2js";

const workbookRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const xmlParseOptions = {
  attrkey: "$",
  explicitArray: false,
  trim: true,
};

export async function extractWorksheetControlMetadataFromWorkbookBuffer(buffer, options = {}) {
  const zip = await JSZip.loadAsync(buffer);
  const workbookDocument = await readXmlPart(zip, "xl/workbook.xml");

  if (!workbookDocument) {
    throw new Error("xl/workbook.xml が見つかりません");
  }

  const workbookRelationships = await readRelationshipMap(zip, "xl/workbook.xml");
  const workbookRoot = getRootNode(workbookDocument, "workbook");
  const sheetsNode = getChildNode(workbookRoot, "sheets");
  const sheetEntries = getChildNodes(sheetsNode, "sheet");
  const worksheets = [];

  for (const sheetEntry of sheetEntries) {
    const relationshipId = getAttributeValue(sheetEntry, "id");

    if (!relationshipId) {
      continue;
    }

    const relationship = workbookRelationships.get(relationshipId);

    if (!relationship || relationship.type !== workbookRelationshipType) {
      continue;
    }

    const worksheetPartPath = resolveTargetPath("xl/workbook.xml", relationship.target);
    const worksheetDocument = await readXmlPart(zip, worksheetPartPath);

    if (!worksheetDocument) {
      continue;
    }

    const worksheetRoot = getRootNode(worksheetDocument, "worksheet");
    const worksheetRelationships = await readRelationshipMap(zip, worksheetPartPath);
    const drawingRelationshipId = getAttributeValue(getChildNode(worksheetRoot, "drawing"), "id");
    const drawingPartPath = drawingRelationshipId
      ? resolveTargetPath(worksheetPartPath, worksheetRelationships.get(drawingRelationshipId)?.target)
      : undefined;
    const drawingShapeNames = drawingPartPath ? await loadDrawingShapeNames(zip, drawingPartPath) : new Map();
    const oleObjectsByShapeId = buildOleObjectMap(getChildNodes(getChildNode(worksheetRoot, "oleObjects"), "oleObject"));
    const controls = [];

    for (const controlNode of getChildNodes(getChildNode(worksheetRoot, "controls"), "control")) {
      const controlRelationshipId = getAttributeValue(controlNode, "id");
      const shapeId = toUnsignedInteger(getAttributeValue(controlNode, "shapeId"));
      const controlPartPath = controlRelationshipId
        ? resolveTargetPath(worksheetPartPath, worksheetRelationships.get(controlRelationshipId)?.target)
        : undefined;
      const classId = controlPartPath ? await loadActiveXClassId(zip, controlPartPath) : null;
      const oleObject = shapeId !== null ? oleObjectsByShapeId.get(String(shapeId)) : undefined;

      controls.push({
        classId,
        codeName: getAttributeValue(controlNode, "name") ?? null,
        progId: oleObject?.progId ?? null,
        shapeId,
        shapeName: shapeId !== null ? drawingShapeNames.get(String(shapeId)) ?? null : null,
      });
    }

    worksheets.push({
      controls,
      sheetCodeName: getAttributeValue(getChildNode(worksheetRoot, "sheetPr"), "codeName") ?? null,
      sheetName: getAttributeValue(sheetEntry, "name") ?? null,
    });
  }

  return {
    version: 1,
    workbook: options.workbookName ?? null,
    worksheets,
  };
}

export async function extractWorksheetControlMetadataFromWorkbookFile(workbookPath) {
  const { readFile } = await import("node:fs/promises");
  const buffer = await readFile(workbookPath);

  return extractWorksheetControlMetadataFromWorkbookBuffer(buffer, {
    workbookName: path.basename(workbookPath),
  });
}

function arrayify(value) {
  if (value === undefined || value === null) {
    return [];
  }

  return Array.isArray(value) ? value : [value];
}

function buildOleObjectMap(oleObjectNodes) {
  const result = new Map();

  for (const oleObjectNode of oleObjectNodes) {
    const shapeId = toUnsignedInteger(getAttributeValue(oleObjectNode, "shapeId"));

    if (shapeId === null) {
      continue;
    }

    result.set(String(shapeId), {
      progId: getAttributeValue(oleObjectNode, "progId") ?? null,
    });
  }

  return result;
}

function getAttributeValue(node, localAttributeName) {
  const attributes = node?.$;

  if (!attributes) {
    return undefined;
  }

  for (const [attributeName, value] of Object.entries(attributes)) {
    if (getLocalName(attributeName).toLowerCase() === localAttributeName.toLowerCase()) {
      return typeof value === "string" ? value : String(value);
    }
  }

  return undefined;
}

function getChildNode(node, localName) {
  return getChildNodes(node, localName)[0];
}

function getChildNodes(node, localName) {
  if (!node || typeof node !== "object") {
    return [];
  }

  const result = [];

  for (const [key, value] of Object.entries(node)) {
    if (key === "$" || getLocalName(key) !== localName) {
      continue;
    }

    result.push(...arrayify(value));
  }

  return result;
}

function getLocalName(qualifiedName) {
  const parts = String(qualifiedName).split(":");
  return parts[parts.length - 1];
}

function getRootNode(document, localName) {
  const rootEntry = Object.entries(document).find(([key]) => getLocalName(key) === localName);

  if (!rootEntry) {
    throw new Error(`${localName} root が見つかりません`);
  }

  return rootEntry[1];
}

async function loadActiveXClassId(zip, partPath) {
  const document = await readXmlPart(zip, partPath);

  if (!document) {
    return null;
  }

  const root = getRootNode(document, "ocx");
  return getAttributeValue(root, "classid") ?? null;
}

async function loadDrawingShapeNames(zip, drawingPartPath) {
  const document = await readXmlPart(zip, drawingPartPath);
  const result = new Map();

  if (!document) {
    return result;
  }

  collectDrawingShapeNames(document, result);
  return result;
}

function collectDrawingShapeNames(node, result) {
  if (Array.isArray(node)) {
    for (const child of node) {
      collectDrawingShapeNames(child, result);
    }
    return;
  }

  if (!node || typeof node !== "object") {
    return;
  }

  for (const [key, value] of Object.entries(node)) {
    if (key === "$") {
      continue;
    }

    if (getLocalName(key) === "cNvPr") {
      for (const childNode of arrayify(value)) {
        const shapeId = toUnsignedInteger(getAttributeValue(childNode, "id"));
        const shapeName = getAttributeValue(childNode, "name");

        if (shapeId !== null && shapeName) {
          result.set(String(shapeId), shapeName);
        }
      }
      continue;
    }

    collectDrawingShapeNames(value, result);
  }
}

async function readRelationshipMap(zip, sourcePartPath) {
  const relationshipsPartPath = getRelationshipsPartPath(sourcePartPath);
  const relationshipsDocument = await readXmlPart(zip, relationshipsPartPath);
  const result = new Map();

  if (!relationshipsDocument) {
    return result;
  }

  const root = getRootNode(relationshipsDocument, "Relationships");

  for (const relationshipNode of getChildNodes(root, "Relationship")) {
    const id = getAttributeValue(relationshipNode, "Id");
    const target = getAttributeValue(relationshipNode, "Target");

    if (!id || !target) {
      continue;
    }

    result.set(id, {
      target,
      type: getAttributeValue(relationshipNode, "Type") ?? null,
    });
  }

  return result;
}

async function readXmlPart(zip, partPath) {
  const zipFile = zip.file(normalizePartPath(partPath));

  if (!zipFile) {
    return undefined;
  }

  const xml = await zipFile.async("string");
  return parseStringPromise(xml, xmlParseOptions);
}

function getRelationshipsPartPath(sourcePartPath) {
  const normalizedSourcePartPath = normalizePartPath(sourcePartPath);
  const directory = path.posix.dirname(normalizedSourcePartPath);
  const basename = path.posix.basename(normalizedSourcePartPath);

  return normalizePartPath(path.posix.join(directory, "_rels", `${basename}.rels`));
}

function normalizePartPath(partPath) {
  return String(partPath).replace(/^\/+/u, "");
}

function resolveTargetPath(sourcePartPath, target) {
  if (!target) {
    return undefined;
  }

  if (String(target).startsWith("/")) {
    return normalizePartPath(target);
  }

  return normalizePartPath(path.posix.join(path.posix.dirname(normalizePartPath(sourcePartPath)), String(target)));
}

function toUnsignedInteger(value) {
  if (value === undefined || value === null || value === "") {
    return null;
  }

  const parsedValue = Number.parseInt(String(value), 10);
  return Number.isNaN(parsedValue) ? null : parsedValue;
}
