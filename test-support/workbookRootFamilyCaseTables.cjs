"use strict";

const worksheetBroadRoot = {
  completion: {
    positive: [
      {
        anchor: 'Worksheets("Sheet One").OLEObjects("CheckBox1").Object.',
        route: "ole",
        scopes: ["extension", "server-worksheet-broad-root-direct"]
      },
      {
        anchor: 'Application.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.',
        route: "ole",
        scopes: ["extension", "server-worksheet-broad-root-direct"]
      },
      {
        anchor: 'Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.',
        route: "ole",
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Application.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.',
        route: "ole",
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.',
        route: "shape",
        scopes: ["extension", "server-worksheet-broad-root-direct"]
      },
      {
        anchor: 'Application.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.',
        route: "shape",
        scopes: ["extension", "server-worksheet-broad-root-direct"]
      },
      {
        anchor: 'Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.',
        route: "shape",
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Application.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.',
        route: "shape",
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.',
        route: "ole",
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.',
        route: "ole",
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Application.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.',
        route: "ole",
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Application.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.',
        route: "ole",
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.',
        route: "shape",
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.',
        route: "shape",
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Application.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.',
        route: "shape",
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Application.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.',
        route: "shape",
        scopes: ["extension", "server-worksheet-broad-root-item"]
      }
    ],
    negative: [
      {
        anchor: 'Sheets("Sheet One").OLEObjects("CheckBox1").Object.',
        reason: "non-target-root",
        scopes: ["extension", "server-worksheet-broad-root-direct"]
      },
      {
        anchor: 'ActiveSheet.OLEObjects("CheckBox1").Object.',
        reason: "non-target-root",
        scopes: ["extension", "server-worksheet-broad-root-direct"]
      },
      {
        anchor: 'Worksheets(1).OLEObjects("CheckBox1").Object.',
        reason: "numeric-selector",
        scopes: ["extension", "server-worksheet-broad-root-direct"]
      },
      {
        anchor: 'Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.',
        reason: "dynamic-selector",
        scopes: ["extension", "server-worksheet-broad-root-direct"]
      },
      {
        anchor: 'Application.Worksheets(1).OLEObjects("CheckBox1").Object.',
        reason: "numeric-selector",
        scopes: ["extension", "server-worksheet-broad-root-direct"]
      },
      {
        anchor: 'Application.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.',
        reason: "dynamic-selector",
        scopes: ["extension", "server-worksheet-broad-root-direct"]
      },
      {
        anchor: 'Worksheets.Item(1).OLEObjects("CheckBox1").Object.',
        reason: "numeric-selector",
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Worksheets.Item(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.',
        reason: "dynamic-selector",
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Application.Worksheets.Item(1).OLEObjects("CheckBox1").Object.',
        reason: "numeric-selector",
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Application.Worksheets.Item(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.',
        reason: "dynamic-selector",
        scopes: ["extension", "server-worksheet-broad-root-item"]
      }
    ]
  },
  hover: {
    positive: [
      {
        anchor: 'Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Valu',
        scopes: ["extension", "server-worksheet-broad-root-direct"]
      },
      {
        anchor: 'Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu',
        scopes: ["extension", "server-worksheet-broad-root-direct"]
      },
      {
        anchor: 'Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.Valu',
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Valu',
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Valu',
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Valu',
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Application.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Valu',
        scopes: ["extension"]
      },
      {
        anchor: 'Application.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Valu',
        scopes: ["extension"]
      },
      {
        anchor: 'Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu',
        scopes: ["extension"]
      },
      {
        anchor: 'Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Valu',
        scopes: ["extension"]
      }
    ],
    negative: [
      {
        anchor: 'Sheets("Sheet One").OLEObjects("CheckBox1").Object.Valu',
        reason: "non-target-root",
        scopes: ["extension"]
      },
      {
        anchor: 'ActiveSheet.OLEObjects("CheckBox1").Object.Valu',
        reason: "non-target-root",
        scopes: ["extension"]
      },
      {
        anchor: 'Worksheets(1).OLEObjects("CheckBox1").Object.Valu',
        reason: "numeric-selector",
        scopes: ["extension"]
      },
      {
        anchor: 'Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Valu',
        reason: "dynamic-selector",
        scopes: ["extension"]
      },
      {
        anchor: 'Worksheets.Item(1).OLEObjects("CheckBox1").Object.Valu',
        reason: "numeric-selector",
        scopes: ["extension"]
      },
      {
        anchor: 'Worksheets.Item(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Valu',
        reason: "dynamic-selector",
        scopes: ["extension"]
      },
      {
        anchor: 'Application.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Valu',
        reason: "numeric-selector",
        scopes: ["extension"]
      },
      {
        anchor: 'Application.Worksheets.Item(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Valu',
        reason: "dynamic-selector",
        scopes: ["extension"]
      }
    ]
  },
  signature: {
    positive: [
      {
        anchor: 'Application.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Select(',
        scopes: ["extension", "server-worksheet-broad-root-direct"]
      },
      {
        anchor: 'Application.Worksheets("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(',
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Application.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(',
        scopes: ["extension", "server-worksheet-broad-root-direct"]
      },
      {
        anchor: 'Application.Worksheets("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(',
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Select(',
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(',
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Application.Worksheets.Item("Sheet One").OLEObjects("CheckBox1").Object.Select(',
        scopes: ["extension"]
      },
      {
        anchor: 'Application.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(',
        scopes: ["extension"]
      },
      {
        anchor: 'Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(',
        scopes: ["extension"]
      },
      {
        anchor: 'Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(',
        scopes: ["extension"]
      },
      {
        anchor: 'Application.Worksheets.Item("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Select(',
        scopes: ["extension", "server-worksheet-broad-root-item"]
      },
      {
        anchor: 'Application.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(',
        scopes: ["extension", "server-worksheet-broad-root-item"]
      }
    ],
    negative: []
  }
};

const applicationWorkbookRoot = {
  completion: {
    positive: [
      {
        anchor: 'Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.',
        route: "ole",
        state: "static",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.',
        route: "ole",
        state: "static",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.',
        route: "shape",
        state: "static",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor: 'Application.ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.',
        route: "shape",
        state: "static",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.',
        route: "ole",
        state: "matched",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.',
        route: "ole",
        state: "matched",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.',
        route: "shape",
        state: "matched",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.',
        route: "shape",
        state: "matched",
        scopes: ["extension", "server-application-shape"]
      }
    ],
    negative: [
      {
        anchor: 'Application.ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.',
        reason: "code-name-selector",
        state: "static",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.ThisWorkbook.Worksheets(1).OLEObjects("CheckBox1").Object.',
        reason: "numeric-selector",
        state: "static",
        scopes: ["extension"]
      },
      {
        anchor: 'Application.ThisWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.',
        reason: "code-name-selector",
        state: "static",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor: 'Application.ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.',
        reason: "code-name-selector",
        state: "static",
        scopes: ["extension"]
      },
      {
        anchor: 'Application.ThisWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.',
        reason: "dynamic-selector",
        state: "static",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.ThisWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.',
        reason: "dynamic-selector",
        state: "static",
        scopes: ["server-application-shape"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.',
        reason: "snapshot-closed",
        state: "closed",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.',
        reason: "snapshot-closed",
        state: "closed",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.',
        reason: "snapshot-closed",
        state: "closed",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.',
        reason: "snapshot-closed",
        state: "closed",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.',
        reason: "code-name-selector",
        state: "matched",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.',
        reason: "code-name-selector",
        state: "matched",
        scopes: ["server-application-shape"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.',
        reason: "dynamic-selector",
        state: "matched",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets(1).Shapes("CheckBox1").OLEFormat.Object.',
        reason: "numeric-selector",
        state: "matched",
        scopes: ["extension"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.',
        reason: "dynamic-selector",
        state: "matched",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor: 'Application.Caller.OLEObjects("CheckBox1").Object.',
        reason: "non-target-root",
        state: "matched",
        scopes: ["extension", "server-application-ole", "server-application-shadowed"]
      },
      {
        anchor: 'Application.Range("A1").Shapes("CheckBox1").OLEFormat.Object.',
        reason: "non-target-root",
        state: "matched",
        scopes: ["extension", "server-application-ole", "server-application-shape", "server-application-shadowed"]
      },
      {
        anchor: 'Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.',
        reason: "shadowed-root",
        state: "shadowed",
        occurrenceIndex: 1,
        scopes: ["extension", "server-application-shadowed"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.',
        reason: "shadowed-root",
        state: "shadowed",
        occurrenceIndex: 1,
        scopes: ["extension", "server-application-shadowed"]
      }
    ]
  },
  hover: {
    positive: [
      {
        anchor: 'Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Valu',
        state: "static",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu',
        state: "static",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Valu',
        state: "matched",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu',
        state: "matched",
        scopes: ["extension", "server-application-shape"]
      }
    ],
    negative: [
      {
        anchor: 'Application.ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Valu',
        reason: "code-name-selector",
        state: "static",
        scopes: ["extension"]
      },
      {
        anchor: 'Application.ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Valu',
        reason: "numeric-selector",
        state: "static",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Valu',
        reason: "code-name-selector",
        state: "static",
        scopes: ["extension"]
      },
      {
        anchor: 'Application.ThisWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Valu',
        reason: "dynamic-selector",
        state: "static",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.ThisWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Valu',
        reason: "numeric-selector",
        state: "static",
        scopes: ["server-application-shape"]
      },
      {
        anchor: 'Application.ThisWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Valu',
        reason: "dynamic-selector",
        state: "static",
        scopes: ["server-application-shape"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Valu',
        reason: "snapshot-closed",
        state: "static",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Valu',
        reason: "snapshot-closed",
        state: "static",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor: 'Application.Caller.OLEObjects("CheckBox1").Object.Valu',
        reason: "non-target-root",
        state: "static",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.Range("A1").Shapes("CheckBox1").OLEFormat.Object.Valu',
        reason: "non-target-root",
        state: "static",
        scopes: ["extension", "server-application-ole", "server-application-shape"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Valu',
        reason: "code-name-selector",
        state: "matched",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Valu',
        reason: "numeric-selector",
        state: "matched",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Valu',
        reason: "dynamic-selector",
        state: "matched",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets(1).Shapes("CheckBox1").OLEFormat.Object.Valu',
        reason: "numeric-selector",
        state: "matched",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Valu',
        reason: "dynamic-selector",
        state: "matched",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor: 'Application.Caller.OLEObjects("CheckBox1").Object.Valu',
        reason: "non-target-root",
        state: "matched",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.Range("A1").Shapes("CheckBox1").OLEFormat.Object.Valu',
        reason: "non-target-root",
        state: "matched",
        scopes: ["extension", "server-application-ole", "server-application-shape"]
      }
    ]
  },
  signature: {
    positive: [
      {
        anchor: 'Application.ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(',
        state: "static",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(',
        state: "static",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(',
        state: "matched",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(',
        state: "matched",
        scopes: ["extension", "server-application-shape"]
      }
    ],
    negative: [
      {
        anchor: 'Application.ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Select(',
        reason: "code-name-selector",
        state: "static",
        scopes: ["extension"]
      },
      {
        anchor: 'Application.ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Select(',
        reason: "numeric-selector",
        state: "static",
        scopes: ["extension"]
      },
      {
        anchor: 'Application.ThisWorkbook.Worksheets.Item("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Select(',
        reason: "code-name-selector",
        state: "static",
        scopes: ["extension"]
      },
      {
        anchor: 'Application.ThisWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Select(',
        reason: "dynamic-selector",
        state: "static",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.ThisWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Select(',
        reason: "dynamic-selector",
        state: "static",
        scopes: ["server-application-shape"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(',
        reason: "snapshot-closed",
        state: "static",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(',
        reason: "snapshot-closed",
        state: "static",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor: 'Application.Caller.OLEObjects("CheckBox1").Object.Select(',
        reason: "non-target-root",
        state: "static",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.Range("A1").Shapes("CheckBox1").OLEFormat.Object.Select(',
        reason: "non-target-root",
        state: "static",
        scopes: ["extension", "server-application-ole", "server-application-shape"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Select(',
        reason: "dynamic-selector",
        state: "matched",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor: 'Application.ActiveWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Select(',
        reason: "dynamic-selector",
        state: "matched",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Application.Caller.OLEObjects("CheckBox1").Object.Select(',
        reason: "non-target-root",
        state: "matched",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Application.Range("A1").Shapes("CheckBox1").OLEFormat.Object.Select(',
        reason: "non-target-root",
        state: "matched",
        scopes: ["extension", "server-application-ole", "server-application-shape"]
      }
    ]
  },
  semantic: {
    positive: [
      {
        anchor: 'Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value',
        identifier: "Value",
        state: "static",
        tokenKind: "property",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor:
          'Call Application.ThisWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(',
        identifier: "Select",
        state: "static",
        tokenKind: "method",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value',
        identifier: "Value",
        state: "static",
        tokenKind: "property",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor:
          'Call Application.ThisWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(',
        identifier: "Select",
        state: "static",
        tokenKind: "method",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor: 'Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value',
        identifier: "Value",
        state: "matched",
        tokenKind: "property",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor:
          'Call Application.ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(',
        identifier: "Select",
        state: "matched",
        tokenKind: "method",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value',
        identifier: "Value",
        state: "matched",
        tokenKind: "property",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor:
          'Call Application.ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(',
        identifier: "Select",
        state: "matched",
        tokenKind: "method",
        scopes: ["extension", "server-application-shape"]
      }
    ],
    negative: [
      {
        anchor: 'Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value',
        identifier: "Value",
        reason: "snapshot-closed",
        state: "static",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor:
          'Call Application.ActiveWorkbook.Worksheets.Item("Sheet One").OLEObjects.Item("CheckBox1").Object.Select(',
        identifier: "Select",
        reason: "snapshot-closed",
        state: "static",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value',
        identifier: "Value",
        reason: "snapshot-closed",
        state: "static",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor:
          'Call Application.ActiveWorkbook.Worksheets.Item("Sheet One").Shapes.Item("CheckBox1").OLEFormat.Object.Select(',
        identifier: "Select",
        reason: "snapshot-closed",
        state: "static",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor: 'Debug.Print Application.ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Value',
        identifier: "Value",
        reason: "code-name-selector",
        state: "static",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Debug.Print Application.ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Value',
        identifier: "Value",
        reason: "numeric-selector",
        state: "static",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Debug.Print Application.ThisWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Value',
        identifier: "Value",
        reason: "dynamic-selector",
        state: "static",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Call Application.ThisWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Select(',
        identifier: "Select",
        reason: "dynamic-selector",
        state: "static",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Debug.Print Application.ActiveWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Value',
        identifier: "Value",
        reason: "code-name-selector",
        state: "static",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Debug.Print Application.ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Value',
        identifier: "Value",
        reason: "numeric-selector",
        state: "static",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Debug.Print Application.ActiveWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Value',
        identifier: "Value",
        reason: "dynamic-selector",
        state: "static",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Call Application.ActiveWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Select(',
        identifier: "Select",
        reason: "dynamic-selector",
        state: "static",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Debug.Print Application.Caller.OLEObjects("CheckBox1").Object.Value',
        identifier: "Value",
        reason: "non-target-root",
        state: "static",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Call Application.Caller.OLEObjects("CheckBox1").Object.Select(',
        identifier: "Select",
        reason: "non-target-root",
        state: "static",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Debug.Print Application.Range("A1").Shapes("CheckBox1").OLEFormat.Object.Value',
        identifier: "Value",
        reason: "non-target-root",
        state: "static",
        scopes: ["server-application-ole", "server-application-shape"]
      },
      {
        anchor: 'Call Application.Range("A1").Shapes("CheckBox1").OLEFormat.Object.Select(',
        identifier: "Select",
        reason: "non-target-root",
        state: "static",
        scopes: ["server-application-ole", "server-application-shape"]
      },
      {
        anchor: 'Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value',
        identifier: "Value",
        reason: "shadowed-root",
        state: "static",
        occurrenceIndex: 1,
        scopes: ["extension"]
      },
      {
        anchor: 'Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value',
        identifier: "Value",
        reason: "shadowed-root",
        state: "static",
        occurrenceIndex: 1,
        scopes: ["extension"]
      },
      {
        anchor: 'Debug.Print Application.ThisWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Value',
        identifier: "Value",
        reason: "code-name-selector",
        state: "static",
        scopes: ["server-application-shape"]
      },
      {
        anchor: 'Debug.Print Application.ThisWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Value',
        identifier: "Value",
        reason: "numeric-selector",
        state: "static",
        scopes: ["server-application-shape"]
      },
      {
        anchor: 'Debug.Print Application.ThisWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Value',
        identifier: "Value",
        reason: "dynamic-selector",
        state: "static",
        scopes: ["server-application-shape"]
      },
      {
        anchor: 'Call Application.ThisWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Select(',
        identifier: "Select",
        reason: "dynamic-selector",
        state: "static",
        scopes: ["server-application-shape"]
      },
      {
        anchor: 'Debug.Print Application.ActiveWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Value',
        identifier: "Value",
        reason: "code-name-selector",
        state: "static",
        scopes: ["server-application-shape"]
      },
      {
        anchor: 'Debug.Print Application.ActiveWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Value',
        identifier: "Value",
        reason: "numeric-selector",
        state: "static",
        scopes: ["server-application-shape"]
      },
      {
        anchor: 'Debug.Print Application.ActiveWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Value',
        identifier: "Value",
        reason: "dynamic-selector",
        state: "static",
        scopes: ["server-application-shape"]
      },
      {
        anchor: 'Call Application.ActiveWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Select(',
        identifier: "Select",
        reason: "dynamic-selector",
        state: "static",
        scopes: ["server-application-shape"]
      },
      {
        anchor: 'Debug.Print Application.ThisWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Value',
        identifier: "Value",
        reason: "code-name-selector",
        state: "matched",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Debug.Print Application.ThisWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Value',
        identifier: "Value",
        reason: "numeric-selector",
        state: "matched",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Debug.Print Application.ThisWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Value',
        identifier: "Value",
        reason: "dynamic-selector",
        state: "matched",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Call Application.ThisWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Select(',
        identifier: "Select",
        reason: "dynamic-selector",
        state: "matched",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Debug.Print Application.ActiveWorkbook.Worksheets("Sheet1").OLEObjects("CheckBox1").Object.Value',
        identifier: "Value",
        reason: "code-name-selector",
        state: "matched",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Debug.Print Application.ActiveWorkbook.Worksheets.Item(1).OLEObjects("CheckBox1").Object.Value',
        identifier: "Value",
        reason: "numeric-selector",
        state: "matched",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Debug.Print Application.ActiveWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Value',
        identifier: "Value",
        reason: "dynamic-selector",
        state: "matched",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Call Application.ActiveWorkbook.Worksheets(GetIndex()).OLEObjects("CheckBox1").Object.Select(',
        identifier: "Select",
        reason: "dynamic-selector",
        state: "matched",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Debug.Print Application.Caller.OLEObjects("CheckBox1").Object.Value',
        identifier: "Value",
        reason: "non-target-root",
        state: "matched",
        scopes: ["extension", "server-application-ole"]
      },
      {
        anchor: 'Call Application.Caller.OLEObjects("CheckBox1").Object.Select(',
        identifier: "Select",
        reason: "non-target-root",
        state: "matched",
        scopes: ["server-application-ole"]
      },
      {
        anchor: 'Debug.Print Application.Range("A1").Shapes("CheckBox1").OLEFormat.Object.Value',
        identifier: "Value",
        reason: "non-target-root",
        state: "matched",
        scopes: ["extension", "server-application-ole", "server-application-shape"]
      },
      {
        anchor: 'Call Application.Range("A1").Shapes("CheckBox1").OLEFormat.Object.Select(',
        identifier: "Select",
        reason: "non-target-root",
        state: "matched",
        scopes: ["server-application-ole", "server-application-shape"]
      },
      {
        anchor: 'Debug.Print Application.ThisWorkbook.Worksheets("Sheet One").OLEObjects("CheckBox1").Object.Value',
        identifier: "Value",
        reason: "shadowed-root",
        state: "matched",
        occurrenceIndex: 1,
        scopes: ["extension"]
      },
      {
        anchor: 'Debug.Print Application.ActiveWorkbook.Worksheets("Sheet One").Shapes("CheckBox1").OLEFormat.Object.Value',
        identifier: "Value",
        reason: "shadowed-root",
        state: "matched",
        occurrenceIndex: 1,
        scopes: ["extension"]
      },
      {
        anchor: 'Debug.Print Application.ThisWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Value',
        identifier: "Value",
        reason: "code-name-selector",
        state: "matched",
        scopes: ["server-application-shape"]
      },
      {
        anchor: 'Debug.Print Application.ThisWorkbook.Worksheets.Item(1).Shapes("CheckBox1").OLEFormat.Object.Value',
        identifier: "Value",
        reason: "numeric-selector",
        state: "matched",
        scopes: ["server-application-shape"]
      },
      {
        anchor: 'Debug.Print Application.ThisWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Value',
        identifier: "Value",
        reason: "dynamic-selector",
        state: "matched",
        scopes: ["server-application-shape"]
      },
      {
        anchor: 'Call Application.ThisWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Select(',
        identifier: "Select",
        reason: "dynamic-selector",
        state: "matched",
        scopes: ["server-application-shape"]
      },
      {
        anchor: 'Debug.Print Application.ActiveWorkbook.Worksheets("Sheet1").Shapes("CheckBox1").OLEFormat.Object.Value',
        identifier: "Value",
        reason: "code-name-selector",
        state: "matched",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor: 'Debug.Print Application.ActiveWorkbook.Worksheets(1).Shapes("CheckBox1").OLEFormat.Object.Value',
        identifier: "Value",
        reason: "numeric-selector",
        state: "matched",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor: 'Debug.Print Application.ActiveWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Value',
        identifier: "Value",
        reason: "dynamic-selector",
        state: "matched",
        scopes: ["extension", "server-application-shape"]
      },
      {
        anchor: 'Call Application.ActiveWorkbook.Worksheets(GetIndex()).Shapes("CheckBox1").OLEFormat.Object.Select(',
        identifier: "Select",
        reason: "dynamic-selector",
        state: "matched",
        scopes: ["server-application-shape"]
      }
    ]
  }
};

module.exports = {
  workbookRootFamilyCaseTables: {
    applicationWorkbookRoot,
    worksheetBroadRoot
  }
};
