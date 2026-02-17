# TODO

## Open Issues

### Inline images not transferring to DevOps
`getAttachmentContentAsync` still throws "message channel closed before a response was received" despite the 200ms delay and 15-second timeout added in the latest build. The `cid:` references in the email body are not getting replaced with DevOps attachment URLs, so images are missing from the work item description. Need to investigate alternative approaches - possibly fetching attachment content via EWS/REST instead of the Office.js callback API, or using `Office.context.mailbox.item.getAttachmentsAsync` to get attachment metadata first.

**Files:** `src/taskpane/services/outlook-mail.js` (getAttachmentContent), `src/taskpane/components/create-ticket.js` (inline image loop starting ~line 228)

### Email body not populating Repro Steps on Issues
When creating an Issue, the email body HTML should be written to `Microsoft.VSTS.TCM.ReproSteps` but it is not appearing on the created work item. Other work item types use `System.Description`. Need to verify the field name is correct for the project's process template and check whether `addAttachmentToWorkItem` is correctly setting the body field value.

**Files:** `src/taskpane/components/create-ticket.js` (getBodyField, handleCreate), `src/taskpane/services/devops-api.js` (addAttachmentToWorkItem)

### Prevent duplicate ticket creation from the same email
Users can click "Create Ticket" multiple times on the same email and create duplicate work items. The create button should be disabled after a successful ticket creation and only re-enabled when the user navigates to a different email. Need to detect email change via `Office.context.mailbox.item` changes or use `addHandlerAsync` with `Office.EventType.ItemChanged`.

**Files:** `src/taskpane/components/create-ticket.js` (handleCreate, init), `src/taskpane/taskpane.js`
