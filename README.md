# CMG Report Builder v0.1

STATUS: source-code prototype. Not hosted, not deployed, not tested in a live Word session.
The included manifest.xml is a TEMPLATE with https://example.invalid/cmg-report-builder.
Uploading that unconfigured manifest will NOT load a working add-in.

## Included
- manifest.xml: XML add-in manifest, CMG Project Tools ribbon tab, Open Report Builder button.
- taskpane.html / taskpane.css / taskpane.js: form and document logic.
- commands.html: ribbon function-file page.
- configure.html: offline manifest generator; enter the approved hosted folder URL.
- preview.html: offline panel visual preview; buttons disabled.
- assets: simple placeholder icons, replace before production.
- fields.json: documented field schema.

## Deployment prerequisite
Host taskpane.html, taskpane.css, taskpane.js, commands.html, help.html and assets on an IT-approved HTTPS web host that permits Office task-pane embedding. A SharePoint document-sharing link is not automatically a compatible web-app host. Do not publish confidential project information or internal documents on a public host. This package contains blank templates, not imported project records.
Open configure.html locally and enter the hosted folder URL to download the configured manifest. Upload the resulting manifest, not the example-address template. HTTPS hosting and tenant approval are not provided by this ZIP.

## First release
Add Project: appends a two-column native Word table, with tagged editable fields.
Refresh Projects: lists cards created by this builder.
Load Project: loads an existing builder card into the form.
Apply Update: updates changed fields only, with a best-effort stale-text check.
Clear Form: clears the panel without changing the document.
Add Department Update: appends an editable department table.
Add Summary Template: inserts a blank director summary at the beginning.

## Boundaries
Does not migrate existing call-sheet sections automatically.
Does not implement report checks, automatic summaries, weekly copying, permissions per project, or identity verification.
Confirmation date is entered/confirmed by the contributor; it is not a verified audit trail.
No Graph permissions, external project database, browser data persistence, telemetry or project-data fetch requests.
Office.js is loaded from Microsoft's hosted CDN. Word handles shared-document persistence.
Native document edits are supported. Use YYYY-MM-DD in date fields and whole numbers without a percent sign in progress fields if using the panel. Reformatting or deleting tagged controls can prevent loading a card.
Concurrency checks are best effort, not atomic locks. Reload before applying updates. Avoid editing the same card at the same time. Test coauthoring behavior in your actual tenant.

## Validation before production
Use a copy of the shared document. Test insertion, loading, direct Word edits, updating, deleted fields, simultaneous edits, page breaks, long updates, PDF export, browser Word, and desktop Word. Keep SharePoint version history enabled according to your team's policy.
This package was checked for XML well-formedness and JavaScript syntax, not Office manifest schema certification or live Office behavior.

## References
https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/word-tutorial
https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests
https://learn.microsoft.com/en-us/javascript/api/manifest/sourcelocation
