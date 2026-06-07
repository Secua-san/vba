var PROVIDER_KIND = "excel-active-workbook";
var VERSION = 1;

function pad2(value) {
  return value < 10 ? "0" + value : String(value);
}

function pad3(value) {
  if (value < 10) {
    return "00" + value;
  }
  if (value < 100) {
    return "0" + value;
  }
  return String(value);
}

function observedAt() {
  var now = new Date();
  return now.getUTCFullYear() + "-" +
    pad2(now.getUTCMonth() + 1) + "-" +
    pad2(now.getUTCDate()) + "T" +
    pad2(now.getUTCHours()) + ":" +
    pad2(now.getUTCMinutes()) + ":" +
    pad2(now.getUTCSeconds()) + "." +
    pad3(now.getUTCMilliseconds()) + "Z";
}

function quote(value) {
  return "\"" + String(value)
    .replace(/\\/g, "\\\\")
    .replace(/"/g, "\\\"")
    .replace(/\r/g, "\\r")
    .replace(/\n/g, "\\n")
    .replace(/\t/g, "\\t") + "\"";
}

function field(name, value) {
  return quote(name) + ":" + value;
}

function baseFields(state) {
  return [
    field("version", String(VERSION)),
    field("providerKind", quote(PROVIDER_KIND)),
    field("state", quote(state)),
    field("observedAt", quote(observedAt()))
  ];
}

function identityFields(workbook) {
  return "{" + [
    field("fullName", quote(workbook.FullName)),
    field("name", quote(workbook.Name)),
    field("path", quote(workbook.Path)),
    field("isAddin", workbook.IsAddin ? "true" : "false")
  ].join(",") + "}";
}

function optionalStringField(fields, name, value) {
  if (value !== null && typeof value !== "undefined") {
    var stringValue = String(value);

    if (stringValue.length > 0 && /\S/.test(stringValue)) {
      fields.push(field(name, quote(stringValue)));
    }
  }
}

function protectedViewFields(protectedWindow) {
  var fields = [];

  try {
    optionalStringField(fields, "sourceName", protectedWindow.SourceName);
  } catch (error) {
  }

  try {
    optionalStringField(fields, "sourcePath", protectedWindow.SourcePath);
  } catch (error) {
  }

  return fields.length > 0 ? "{" + fields.join(",") + "}" : "";
}

function emitSnapshot(fields) {
  WScript.StdOut.Write("{" + fields.join(",") + "}");
}

function emitUnavailable(reason) {
  var fields = baseFields("unavailable");
  fields.push(field("reason", quote(reason)));
  emitSnapshot(fields);
}

function emitProtectedView(protectedWindow) {
  var fields = baseFields("protected-view");
  var protectedView = protectedViewFields(protectedWindow);

  if (protectedView.length > 0) {
    fields.push(field("protectedView", protectedView));
  }

  emitSnapshot(fields);
}

function emitUnsupported(reason, workbook) {
  var fields = baseFields("unsupported");
  fields.push(field("reason", quote(reason)));
  fields.push(field("identity", identityFields(workbook)));
  emitSnapshot(fields);
}

function emitAvailable(workbook) {
  var fields = baseFields("available");
  fields.push(field("identity", identityFields(workbook)));
  emitSnapshot(fields);
}

function getActiveProtectedViewWindow(application) {
  try {
    var protectedWindow = application.ActiveProtectedViewWindow;
    return protectedWindow !== null && typeof protectedWindow !== "undefined" ? protectedWindow : null;
  } catch (error) {
    return null;
  }
}

function main() {
  var application;

  try {
    application = GetObject("", "Excel.Application");
  } catch (error) {
    emitUnavailable("host-unreachable");
    return;
  }

  var protectedWindow = getActiveProtectedViewWindow(application);
  if (protectedWindow !== null) {
    emitProtectedView(protectedWindow);
    return;
  }

  var workbook = application.ActiveWorkbook;
  if (workbook === null || typeof workbook === "undefined") {
    emitUnavailable("no-active-workbook");
    return;
  }

  if (workbook.IsAddin) {
    emitUnsupported("addin", workbook);
    return;
  }

  if (String(workbook.Path).length === 0) {
    emitUnsupported("unsaved", workbook);
    return;
  }

  emitAvailable(workbook);
}

try {
  main();
} catch (error) {
  emitUnavailable("host-error");
}
