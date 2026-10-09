"use strict";

// key, label shown in the Word table, panel input id
const fields = [
  ["site", "Site"],
  ["engineer", "Project Engineer Name"],
  ["start", "Start Date"],
  ["contractEnd", "Contract End Date"],
  ["progress", "Percent Complete"],
  ["status", "Project Status"],
  ["tasks", "Remaining Tasks"]
];
const requiredFields = ["site", "engineer"];

let loadedId = null;
let baseline = {};
let busy = false;

const $ = id => document.getElementById(id);
const clean = s => String(s || "").replace(/\u0007/g, "").replace(/\r/g, "\n").trim();
const message = text => { $("message").textContent = text; };
const readForm = () => Object.fromEntries(fields.map(([k]) => [k, $(k).value.trim()]));

function validate(d) {
  const missing = requiredFields.filter(k => !d[k]);
  if (missing.length) throw new Error("Enter Site and Project Engineer Name first.");
}

async function action(fn) {
  if (busy) return;
  busy = true;
  document.querySelectorAll("button").forEach(b => b.disabled = true);
  try {
    await fn();
  } catch (e) {
    message("Action stopped: " + (e.message || String(e)) + " Check the document before retrying if an insertion partially completed.");
  } finally {
    busy = false;
    document.querySelectorAll("button").forEach(b => b.disabled = false);
    $("update").disabled = !loadedId;
  }
}

function clearForm() {
  fields.forEach(([k]) => { $(k).value = ""; });
  loadedId = null;
  baseline = {};
  $("mode").textContent = "New project mode";
  $("update").disabled = true;
}

// ---------- Project card layout ----------
// Row 1 (header): Site | empty | Project Engineer Name
// Row 2: Start Date | Contract End Date | Percent Complete
// Row 3: Project Status | content (columns 2 and 3 merged)
// Row 4: Remaining Tasks | content (columns 2 and 3 merged)
const BROWN = "#351C15";
const LIGHT = "#F2F0ED";
const LINE = "#C9C4BF";

// Writes "Label: value" in one cell. Only the value is a content control.
function fillLabelValue(cell, label, value) {
  const p = cell.body.paragraphs.getFirst();
  const l = p.insertText(label + ": ", "Start");
  l.font.bold = true;
  const v = p.insertText(value || " ", "End");
  v.font.bold = false;
  return v.insertContentControl();
}

// Writes only a value (used in the merged cells).
function fillValue(cell, value) {
  const p = cell.body.paragraphs.getFirst();
  const v = p.insertText(value || " ", "Start");
  v.font.bold = false;
  return v.insertContentControl();
}

function styleCard(table) {
  table.font.name = "Aptos";
  table.font.size = 10;
  table.autoFitWindow();
  table.setCellPadding("Top", 3);
  table.setCellPadding("Bottom", 3);
  table.setCellPadding("Left", 5);
  table.setCellPadding("Right", 5);
  const all = table.getBorder("All");
  all.type = "Single";
  all.color = LINE;
  all.width = 0.5;

  // Header row: brown, white text
  for (let c = 0; c < 3; c++) {
    const cell = table.getCell(0, c);
    cell.shadingColor = BROWN;
    cell.body.font.color = "#FFFFFF";
  }
  // Empty header column: no left and right border (and match its neighbors)
  table.getCell(0, 1).getBorder("Left").type = "None";
  table.getCell(0, 1).getBorder("Right").type = "None";
  table.getCell(0, 0).getBorder("Right").type = "None";
  table.getCell(0, 2).getBorder("Left").type = "None";

  // Label cells in rows 3 and 4
  for (let r = 2; r <= 3; r++) {
    const cell = table.getCell(r, 0);
    cell.shadingColor = LIGHT;
    cell.body.font.bold = true;
    cell.body.font.color = BROWN;
  }
}

async function insertProject() {
  if (loadedId) throw new Error("Clear Form before adding another project. Use Apply Update for the loaded card.");
  const d = readForm();
  validate(d);
  const id = "cmg-project-" + (crypto.randomUUID ? crypto.randomUUID() : Date.now() + "-" + Math.random().toString(16).slice(2));

  await Word.run(async context => {
    const body = context.document.body;
    body.insertParagraph("", Word.InsertLocation.end);

    // Table is inserted in the text flow (inline), with no text wrapping settings.
    const table = body.insertTable(4, 3, Word.InsertLocation.end, [
      ["", "", ""],
      ["", "", ""],
      ["Project Status:", "", ""],
      ["Remaining Tasks:", "", ""]
    ]);
    table.mergeCells(2, 1, 2, 2);
    table.mergeCells(3, 1, 3, 2);
    await context.sync();

    styleCard(table);

    const label = Object.fromEntries(fields);
    const controls = {
      site: fillLabelValue(table.getCell(0, 0), label.site, d.site),
      engineer: fillLabelValue(table.getCell(0, 2), label.engineer, d.engineer),
      start: fillLabelValue(table.getCell(1, 0), label.start, d.start),
      contractEnd: fillLabelValue(table.getCell(1, 1), label.contractEnd, d.contractEnd),
      progress: fillLabelValue(table.getCell(1, 2), label.progress, d.progress),
      status: fillValue(table.getCell(2, 1), d.status),
      tasks: fillValue(table.getCell(3, 1), d.tasks)
    };
    for (const [k, cc] of Object.entries(controls)) {
      cc.tag = id + ":" + k;
      cc.title = label[k];
    }
    // Keep header values white on the brown background
    ["site", "engineer"].forEach(k => { controls[k].font.color = "#FFFFFF"; });

    const card = table.getRange().insertContentControl();
    card.tag = id;
    card.title = "CMG Project Update";
    body.insertParagraph("", Word.InsertLocation.end);
    await context.sync();
  });
  clearForm();
  await refreshProjects();
  message("Project card added at the end of the document.");
}

async function refreshProjects() {
  await Word.run(async context => {
    const controls = context.document.contentControls;
    controls.load("items/id,items/tag,items/title");
    await context.sync();
    const cards = controls.items.filter(c => /^cmg-project-[^:]+$/.test(c.tag));
    const sites = cards.map(c => ({ id: c.id, field: controls.items.find(f => f.tag === c.tag + ":site") }));
    sites.forEach(s => { if (s.field) s.field.load("text"); });
    await context.sync();
    $("projects").replaceChildren(new Option("Select a project", ""));
    sites.forEach(s => $("projects").add(new Option(s.field ? clean(s.field.text) || "Untitled project" : "Project card", String(s.id))));
  });
}

async function loadProject() {
  loadedId = null;
  baseline = {};
  const id = Number($("projects").value);
  if (!id) throw new Error("Select a project from the dropdown.");
  await Word.run(async context => {
    const card = context.document.contentControls.getById(id);
    card.load("tag");
    const controls = card.contentControls;
    controls.load("items/tag,items/text");
    await context.sync();
    fields.forEach(([k]) => {
      const c = controls.items.find(c => c.tag === card.tag + ":" + k);
      if (!c) throw new Error("This card is missing the " + k + " field. Edit it directly in Word instead.");
      baseline[k] = clean(c.text);
      $(k).value = baseline[k];
    });
    loadedId = id;
  });
  $("mode").textContent = "Editing loaded project: " + $("site").value;
  message("Project loaded. Change fields, then select Apply Update.");
}

async function updateProject() {
  if (!loadedId) throw new Error("Load a project first.");
  const d = readForm();
  validate(d);
  const changed = fields.filter(([k]) => d[k] !== baseline[k]);
  if (!changed.length) { message("No changes to apply."); return; }
  await Word.run(async context => {
    const card = context.document.contentControls.getById(loadedId);
    card.load("tag");
    const controls = card.contentControls;
    controls.load("items/tag,items/text");
    await context.sync();
    const targets = changed.map(([k]) => ({ k, c: controls.items.find(c => c.tag === card.tag + ":" + k) }));
    for (const { k, c } of targets) {
      if (!c) throw new Error("A required field was removed. Reload the card.");
      if (clean(c.text) !== baseline[k]) throw new Error("The " + k + " field changed in Word since you loaded it. Reload before applying your update.");
    }
    for (const { k, c } of targets) c.insertText(d[k] || " ", Word.InsertLocation.replace);
    await context.sync();
  });
  baseline = { ...d };
  await refreshProjects();
  message("Changed fields updated. Word handles saving the shared document.");
}

// ---------- Department update and director summary (unchanged) ----------
function styleTable(table, count) {
  table.font.name = "Aptos";
  table.font.size = 10;
  table.styleBuiltIn = Word.BuiltInStyleName.gridTable1Light_Accent1;
  table.styleFirstColumn = false;
  table.getCell(0, 0).shadingColor = "#351C15";
  table.getCell(0, 1).shadingColor = "#351C15";
  table.getCell(0, 0).body.font.color = "#FFFFFF";
  table.getCell(0, 1).body.font.color = "#FFB500";
  table.getCell(0, 0).body.font.bold = true;
  table.getCell(0, 1).body.font.bold = true;
  for (let i = 1; i < count; i++) {
    table.getCell(i, 0).shadingColor = "#F2F0ED";
    table.getCell(i, 0).body.font.bold = true;
  }
}

async function insertSimple(title, rows, location) {
  await Word.run(async context => {
    const body = context.document.body;
    const table = body.insertTable(rows.length + 1, 2, location, [[title, "CMG WEEKLY REPORT"], ...rows]);
    styleTable(table, rows.length + 1);
    table.getRange().insertContentControl().title = title;
    if (location === Word.InsertLocation.end) body.insertParagraph("", Word.InsertLocation.end);
    await context.sync();
  });
  message(title + " inserted. Edit its cells directly in Word.");
}

if (typeof Office !== "undefined") Office.onReady(async info => {
  if (info.host !== Office.HostType.Word) { message("Open this panel through the Word add-in. For a visual preview, use preview.html."); return; }
  if (!Office.context.requirements.isSetSupported("WordApi", "1.4")) { message("This version needs WordApi 1.4 (merged table cells). Update Word or use a supported browser version."); return; }
  $("insert").onclick = () => action(insertProject);
  $("refresh").onclick = () => action(async () => { await refreshProjects(); message("Project list refreshed."); });
  $("load").onclick = () => action(loadProject);
  $("new").onclick = () => { clearForm(); message("Form cleared. Existing document content was not changed."); };
  $("update").onclick = () => action(updateProject);
  $("department").onclick = () => action(async () => {
    if (!$("topic").value.trim()) throw new Error("Enter a department topic first.");
    await insertSimple("DEPARTMENT UPDATE", [["Topic", $("topic").value], ["Owner", $("deptOwner").value], ["This Week", $("deptWeekly").value], ["Next Action", $("deptNext").value], ["Help Needed", $("deptHelp").value]], Word.InsertLocation.end);
  });
  $("summary").onclick = () => action(() => insertSimple("DIRECTOR SUMMARY", [["Meeting Date", ""], ["Decisions Needed", ""], ["Projects Needing Attention", ""], ["Upcoming Milestones", ""], ["Work Complete / Closeout Pending", ""]], Word.InsertLocation.start));
  await action(async () => { await refreshProjects(); message("Ready. Add a project or load an existing builder card."); });
});
