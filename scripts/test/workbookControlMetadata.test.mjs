import assert from "node:assert/strict";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import test from "node:test";

import JSZip from "jszip";

import {
  extractWorksheetControlMetadataFromWorkbookBuffer,
  extractWorksheetControlMetadataFromWorkbookFile,
} from "../lib/workbookControlMetadata.mjs";

const execFileAsync = promisify(execFile);

test("worksheet workbook package から shape name / code name / ProgID / classId を抽出する", async () => {
  const workbookBuffer = await createWorkbookPackageBuffer();
  const metadata = await extractWorksheetControlMetadataFromWorkbookBuffer(workbookBuffer);

  assert.deepEqual(metadata, {
    version: 1,
    workbook: null,
    worksheets: [
      {
        controls: [
          {
            classId: "{8BD21D40-EC42-11CE-9E0D-00AA006002F3}",
            codeName: "chkFinished",
            progId: "Forms.CheckBox.1",
            shapeId: 3,
            shapeName: "CheckBox1",
          },
        ],
        sheetCodeName: "Sheet1",
        sheetName: "Sheet1",
      },
    ],
  });
});

test("chartsheet は worksheet probe の対象外として無視する", async () => {
  const workbookBuffer = await createWorkbookPackageBuffer({ includeChartsheet: true });
  const metadata = await extractWorksheetControlMetadataFromWorkbookBuffer(workbookBuffer);

  assert.equal(metadata.worksheets.length, 1);
  assert.equal(metadata.worksheets[0]?.sheetName, "Sheet1");
});

test("CLI は workbook 名付きの JSON を出力する", async () => {
  const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "vba-workbook-control-metadata-"));
  const workbookPath = path.join(temporaryDirectory, "fixture.xlsm");

  try {
    await writeFile(workbookPath, await createWorkbookPackageBuffer());

    const { stdout } = await execFileAsync(process.execPath, [
      path.resolve("scripts", "probe-workbook-control-metadata.mjs"),
      workbookPath,
    ], {
      cwd: path.resolve("."),
    });

    const metadata = JSON.parse(stdout);

    assert.equal(metadata.workbook, "fixture.xlsm");
    assert.equal(metadata.worksheets[0]?.sheetCodeName, "Sheet1");
    assert.equal(metadata.worksheets[0]?.controls[0]?.shapeName, "CheckBox1");
    assert.equal(metadata.worksheets[0]?.controls[0]?.codeName, "chkFinished");
  } finally {
    await rm(temporaryDirectory, { force: true, recursive: true });
  }
});

test("file helper は workbook 名を保持する", async () => {
  const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), "vba-workbook-control-metadata-"));
  const workbookPath = path.join(temporaryDirectory, "fixture.xlam");

  try {
    await writeFile(workbookPath, await createWorkbookPackageBuffer());

    const metadata = await extractWorksheetControlMetadataFromWorkbookFile(workbookPath);

    assert.equal(metadata.workbook, "fixture.xlam");
    assert.equal(metadata.worksheets[0]?.controls[0]?.classId, "{8BD21D40-EC42-11CE-9E0D-00AA006002F3}");
  } finally {
    await rm(temporaryDirectory, { force: true, recursive: true });
  }
});

async function createWorkbookPackageBuffer(options = {}) {
  const zip = new JSZip();
  const workbookRelationships = [
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>`,
  ];
  const workbookSheets = [
    `<sheet name="Sheet1" sheetId="1" r:id="rId1"/>`,
  ];

  if (options.includeChartsheet) {
    workbookRelationships.push(
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet2.xml"/>`,
    );
    workbookSheets.push(`<sheet name="Chart1" sheetId="2" r:id="rId2"/>`);
  }

  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    ${workbookSheets.join("\n    ")}
  </sheets>
</workbook>`,
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${workbookRelationships.join("\n  ")}
</Relationships>`,
  );
  zip.file(
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetPr codeName="Sheet1" />
  <drawing r:id="rId1" />
  <controls>
    <control r:id="rId2" shapeId="3" name="chkFinished" />
  </controls>
  <oleObjects>
    <oleObject progId="Forms.CheckBox.1" shapeId="3" />
  </oleObjects>
</worksheet>`,
  );
  zip.file(
    "xl/worksheets/_rels/sheet1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/>
</Relationships>`,
  );
  zip.file(
    "xl/drawings/drawing1.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
  <xdr:twoCellAnchor>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="3" name="CheckBox1" />
        <xdr:cNvSpPr />
      </xdr:nvSpPr>
      <xdr:spPr />
    </xdr:sp>
    <xdr:clientData />
  </xdr:twoCellAnchor>
</xdr:wsDr>`,
  );
  zip.file(
    "xl/ctrlProps/ctrlProp1.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" ax:classid="{8BD21D40-EC42-11CE-9E0D-00AA006002F3}" />`,
  );

  if (options.includeChartsheet) {
    zip.file(
      "xl/chartsheets/sheet2.xml",
      `<?xml version="1.0" encoding="UTF-8"?>
<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetPr codeName="Chart1" />
  <drawing r:id="rId1" />
</chartsheet>`,
    );
  }

  return zip.generateAsync({ type: "nodebuffer" });
}
