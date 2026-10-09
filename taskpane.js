"use strict";
const fields = [["site", "Site / Building Code", "text"], ["owner", "Project Owner", "text"], ["contractor", "Contractor", "text"], ["ownership", "Owned or Leased", "select"], ["status", "Status", "select"], ["progress", "Percent Complete", "number"], ["start", "Start Date", "date"], ["contractEnd", "Contract End Date", "date"], ["forecast", "Forecast Completion", "date"], ["weekly", "This Week\u2019s Update", "textarea"], ["next", "Next Action / Due Date", "textarea"], ["issues", "Risks / Open Issues", "textarea"], ["decision", "Director Decision Needed", "textarea"], ["closeout", "Closeout / Payment Update", "textarea"], ["confirmed", "Update Confirmed On", "date"]];
let loadedId = null;
let baseline = {};
let busy = false;
const $ = id => document.getElementById(id);
const clean = s => String(s || "").replace(/[\r\u0007]/g, "").trim();
const message = text => { $("message").textContent = text; };
const readForm = () => Object.fromEntries(fields.map(([k]) => [k, $(k).value.trim()]));
function validate(d) {
  if (!d.site || !d.owner) throw new Error("Enter Site and Project Owner first.");
  if (d.progress !== "" && (!/^\d+$/.test(d.progress) || Number(d.progress) > 100)) throw new Error("Percent Complete must be a whole number from 0 to 100.");
}
function today() {
  const d = new Date();
  return [d.getFullYear(), String(d.getMonth()+1).padStart(2,"0"), String(d.getDate()).padStart(2,"0")].join("-");
}
async function action(fn) {
  if (busy) return;
  busy = true;
  document.querySelectorAll("button").forEach(b => b.disabled = true);
  try { await fn(); } catch(e) { message("Action stopped: " + (e.message || String(e)) + " Check the document before retrying if an insertion partially completed."); }
  finally { busy = false; document.querySelectorAll("button").forEach(b => b.disabled = false); $("update").disabled = !loadedId; }
}
function clearForm() {
  fields.forEach(([k]) => $(k).value = "");
  loadedId = null; baseline = {};
  $("mode").textContent = "New project mode";
  $("update").disabled = true;
}
function styleTable(table, count) {
  table.font.name = "Aptos";
  table.font.size = 10;
  table.styleBuiltIn = Word.BuiltInStyleName.gridTable1Light_Accent1;
  table.styleFirstColumn = false;
  table.getCell(0,0).shadingColor = "#351C15";
  table.getCell(0,1).shadingColor = "#351C15";
  table.getCell(0,0).body.font.color = "#FFFFFF";
  table.getCell(0,1).body.font.color = "#FFB500";
  table.getCell(0,0).body.font.bold = true;
  table.getCell(0,1).body.font.bold = true;
  for (let i=1;i<count;i++) {
    table.getCell(i,0).shadingColor = "#F2F0ED";
    table.getCell(i,0).body.font.bold = true;
  }
}
async function insertProject() {
  if (loadedId) throw new Error("Clear Form before adding another project. Use Apply Update for the loaded card.");
  const d = readForm(); validate(d);
  if (!d.confirmed) d.confirmed = today();
  const id = "cmg-project-" + (crypto.randomUUID ? crypto.randomUUID() : Date.now()+"-"+Math.random().toString(16).slice(2));
  await Word.run(async context => {
    const body = context.document.body;
    body.insertParagraph("", Word.InsertLocation.end);
    const values = [["PROJECT UPDATE", "CMG WEEKLY REPORT"], ...fields.map(([k,l]) => [l,d[k] || " "])];
    const table = body.insertTable(values.length,2,Word.InsertLocation.end,values);
    styleTable(table, values.length);
    const card = table.getRange().insertContentControl();
    card.tag = id; card.title = "CMG Project Update";
    for (let i=0;i<fields.length;i++) {
      const [k,l] = fields[i];
      const control = table.getCell(i+1,1).body.getRange(Word.RangeLocation.content).insertContentControl();
      control.tag = id + ":" + k;
      control.title = l;

    }
    body.insertParagraph("",Word.InsertLocation.end);
    await context.sync();
  });
  clearForm(); await refreshProjects(); message("Project card added at the end of the document.");
}
async function refreshProjects() {
  await Word.run(async context => {
    const controls = context.document.contentControls;
    controls.load("items/id,items/tag,items/title");
    await context.sync();
    const cards = controls.items.filter(c => /^cmg-project-[^:]+$/.test(c.tag));
    const sites = cards.map(c => ({id:c.id, field:controls.items.find(f => f.tag===c.tag+":site")}));
    sites.forEach(s => { if(s.field) s.field.load("text"); });
    await context.sync();
    $("projects").replaceChildren(new Option("Select a project",""));
    sites.forEach(s => $("projects").add(new Option(s.field ? clean(s.field.text) || "Untitled project" : "Project card",String(s.id))));
  });
}
async function loadProject() {
  loadedId = null; baseline = {};
  const id = Number($("projects").value);
  if (!id) throw new Error("Select a project from the dropdown.");
  await Word.run(async context => {
    const card = context.document.contentControls.getById(id);
    card.load("tag");
    const controls = card.contentControls;
    controls.load("items/tag,items/text");
    await context.sync();
    baseline = {};
    fields.forEach(([k]) => {
      const c = controls.items.find(c => c.tag===card.tag+":"+k);
      if (!c) throw new Error("This card is missing the " + k + " field. Edit it directly in Word instead.");
      baseline[k] = clean(c.text);
      const input = $(k);
      if(input.tagName === "SELECT" && !Array.from(input.options).some(o=>o.value===baseline[k])) input.add(new Option(baseline[k],baseline[k]));
      input.value = baseline[k];
      if (input.value !== baseline[k]) throw new Error("The " + k + " field was edited into a format the panel cannot load. Correct it directly in Word first.");
    });
    loadedId = id;
  });
  $("mode").textContent = "Editing loaded project: " + $("site").value;
  message("Project loaded. Change fields, then select Apply Update.");
}
async function updateProject() {
  if (!loadedId) throw new Error("Load a project first.");
  const d=readForm(); validate(d);
  const changed=fields.filter(([k])=>d[k]!==baseline[k]);
  if(!changed.length) { message("No changes to apply. Change Update Confirmed On to confirm an unchanged update."); return; }
  await Word.run(async context => {
    const card=context.document.contentControls.getById(loadedId);
    card.load("tag");
    const controls=card.contentControls;
    controls.load("items/tag,items/text");
    await context.sync();
    const targets=changed.map(([k])=>({k,c:controls.items.find(c=>c.tag===card.tag+":"+k)}));
    for(const {k,c} of targets) {
      if(!c) throw new Error("A required field was removed. Reload the card.");
      if(clean(c.text)!==baseline[k]) throw new Error("The "+k+" field changed in Word since you loaded it. Reload before applying your update.");
    }
    for(const {k,c} of targets) c.insertText(d[k],Word.InsertLocation.replace);
    await context.sync();
  });
  baseline={...d}; await refreshProjects(); message("Changed fields updated. Word handles saving the shared document.");
}
async function insertSimple(title,rows,location) {
  await Word.run(async context=>{
    const body=context.document.body;
    const table=body.insertTable(rows.length+1,2,location,[[title,"CMG WEEKLY REPORT"],...rows]);
    styleTable(table,rows.length+1);
    table.getRange().insertContentControl().title=title;
    if(location===Word.InsertLocation.end) body.insertParagraph("",Word.InsertLocation.end);
    await context.sync();
  });
  message(title+" inserted. Edit its cells directly in Word.");
}
if(typeof Office !== "undefined") Office.onReady(async info=>{
  if(info.host!==Office.HostType.Word) {message("Open this panel through the Word add-in. For a visual preview, use preview.html.");return;}
  if(!Office.context.requirements.isSetSupported("WordApi","1.3")) {message("This prototype requires WordApi 1.3. Update Word or use a supported browser version.");return;}
  $("insert").onclick=()=>action(insertProject);
  $("refresh").onclick=()=>action(async()=>{await refreshProjects();message("Project list refreshed.");});
  $("load").onclick=()=>action(loadProject);
  $("new").onclick=()=>{clearForm();message("Form cleared. Existing document content was not changed.");};
  $("update").onclick=()=>action(updateProject);
  $("department").onclick=()=>action(async()=>{
    if(!$("topic").value.trim()) throw new Error("Enter a department topic first.");
    await insertSimple("DEPARTMENT UPDATE",[["Topic",$("topic").value],["Owner",$("deptOwner").value],["This Week",$("deptWeekly").value],["Next Action",$("deptNext").value],["Help Needed",$("deptHelp").value]],Word.InsertLocation.end);
  });
  $("summary").onclick=()=>action(()=>insertSimple("DIRECTOR SUMMARY",[["Meeting Date",""],["Decisions Needed",""],["Projects Needing Attention",""],["Upcoming Milestones",""],["Work Complete / Closeout Pending",""]],Word.InsertLocation.start));
  await action(async()=>{await refreshProjects();message("Ready. Add a project or load an existing builder card.");});
});
