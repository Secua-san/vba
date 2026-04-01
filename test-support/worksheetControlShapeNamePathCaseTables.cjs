"use strict";

const OLE_FIXTURE = "packages/extension/test/fixtures/OleObjectBuiltIn.bas";
const SHAPE_FIXTURE = "packages/extension/test/fixtures/ShapesBuiltIn.bas";

const worksheetControlShapeNamePath = {
  completion: {
    positive: [
      {
        fixture: OLE_FIXTURE,
        anchor: 'Sheet1.OLEObjects("CheckBox1").Object.',
        rootKind: "document-module",
        routeKind: "ole-object",
        scopes: ["extension", "server-worksheet-control-shape-name-path-ole"]
      },
      {
        fixture: OLE_FIXTURE,
        anchor: 'ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.',
        rootKind: "workbook-qualified-static",
        routeKind: "ole-object",
        scopes: ["extension", "server-worksheet-control-shape-name-path-ole"]
      },
      {
        fixture: OLE_FIXTURE,
        anchor: 'ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.',
        rootKind: "workbook-qualified-matched",
        routeKind: "ole-object",
        scopes: ["extension", "server-worksheet-control-shape-name-path-ole"]
      },
      {
        fixture: SHAPE_FIXTURE,
        anchor: 'Sheet1.Shapes("CheckBox1").OLEFormat.Object.',
        rootKind: "document-module",
        routeKind: "shape-oleformat",
        scopes: ["extension", "server-worksheet-control-shape-name-path-shape"]
      },
      {
        fixture: SHAPE_FIXTURE,
        anchor: 'ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.',
        rootKind: "workbook-qualified-static",
        routeKind: "shape-oleformat",
        scopes: ["extension", "server-worksheet-control-shape-name-path-shape"]
      },
      {
        fixture: SHAPE_FIXTURE,
        anchor: 'ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.',
        rootKind: "workbook-qualified-matched",
        routeKind: "shape-oleformat",
        scopes: ["extension", "server-worksheet-control-shape-name-path-shape"]
      }
    ],
    negative: [
      {
        fixture: OLE_FIXTURE,
        anchor: "Sheet1.OLEObjects(i + 1).Object.",
        reason: "dynamic-selector",
        rootKind: "document-module",
        routeKind: "ole-object",
        scopes: ["extension", "server-worksheet-control-shape-name-path-ole"]
      },
      {
        fixture: OLE_FIXTURE,
        anchor: 'Chart1.OLEObjects("CheckBox1").Object.',
        reason: "chartsheet-root",
        rootKind: "document-module",
        routeKind: "ole-object",
        scopes: ["extension", "server-worksheet-control-shape-name-path-ole"]
      },
      {
        fixture: OLE_FIXTURE,
        anchor: 'ActiveSheet.OLEObjects("CheckBox1").Object.',
        reason: "non-target-root",
        rootKind: "document-module",
        routeKind: "ole-object",
        scopes: ["extension", "server-worksheet-control-shape-name-path-ole"]
      },
      {
        fixture: OLE_FIXTURE,
        anchor: 'ThisWorkbook.Worksheets(1).OLEObjects("CheckBox1").Object.',
        reason: "numeric-selector",
        rootKind: "workbook-qualified-static",
        routeKind: "ole-object",
        scopes: ["extension", "server-worksheet-control-shape-name-path-ole"]
      },
      {
        fixture: OLE_FIXTURE,
        anchor: 'ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.',
        reason: "code-name-selector",
        rootKind: "workbook-qualified-static",
        routeKind: "ole-object",
        scopes: ["extension", "server-worksheet-control-shape-name-path-ole"]
      },
      {
        fixture: OLE_FIXTURE,
        anchor: 'ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.',
        rootKind: "workbook-qualified-closed",
        routeKind: "ole-object",
        scopes: ["extension", "server-worksheet-control-shape-name-path-ole"]
      },
      {
        fixture: SHAPE_FIXTURE,
        anchor: 'Chart1.Shapes("CheckBox1").OLEFormat.Object.',
        reason: "chartsheet-root",
        rootKind: "document-module",
        routeKind: "shape-oleformat",
        scopes: ["extension", "server-worksheet-control-shape-name-path-shape"]
      },
      {
        fixture: SHAPE_FIXTURE,
        anchor: 'Sheet1.Shapes("PlainShape").OLEFormat.Object.',
        reason: "plain-shape",
        rootKind: "document-module",
        routeKind: "shape-oleformat",
        scopes: ["extension", "server-worksheet-control-shape-name-path-shape"]
      },
      {
        fixture: SHAPE_FIXTURE,
        anchor: 'ThisWorkbook.Worksheets(1).Shapes("CheckBox1").OLEFormat.Object.',
        reason: "numeric-selector",
        rootKind: "workbook-qualified-static",
        routeKind: "shape-oleformat",
        scopes: ["extension", "server-worksheet-control-shape-name-path-shape"]
      },
      {
        fixture: SHAPE_FIXTURE,
        anchor: 'ThisWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.',
        reason: "code-name-selector",
        rootKind: "workbook-qualified-static",
        routeKind: "shape-oleformat",
        scopes: ["extension", "server-worksheet-control-shape-name-path-shape"]
      },
      {
        fixture: SHAPE_FIXTURE,
        anchor: 'ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.',
        rootKind: "workbook-qualified-closed",
        routeKind: "shape-oleformat",
        scopes: ["extension", "server-worksheet-control-shape-name-path-shape"]
      }
    ]
  }
};

module.exports = {
  worksheetControlShapeNamePathCaseTables: {
    worksheetControlShapeNamePath
  }
};
