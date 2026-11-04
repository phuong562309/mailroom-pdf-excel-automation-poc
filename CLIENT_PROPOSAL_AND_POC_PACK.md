# Mailroom PDF automation: client proposal and proof pack

Prepared 4 November 2026. Drafts for the owner to send; nothing has been posted or submitted.

## Evidence from this proof of concept

The attached synthetic sample contains **15 documents across 24 pages and all 11 placeholder types** (`Type_01`–`Type_11`). It includes one-, two- and three-page documents, tables, Unicode names, and the same person with separate records/statuses. The pipeline produces 15 child PDFs, a two-sheet Excel workbook, and a local `SharePoint_Synced_Outbox/<Document_Type>/` tree.

Independent acceptance: **15/15 documents and 24/24 pages checked; unchecked=0, mismatched=0; 11/11 types**. See `AUDIT.json`, `PORTABILITY_REPORT.json` and `completion.json` for the actual run, hashes, duration and test result. This statement is valid only with the matching accepted audit shipped in this package.

The Maker uses pypdf. The independent Checker reads the original and every child PDF with pdfplumber/pdfminer and reads back the saved Excel workbook. It compares page coverage/order/content, identifiers, names, statuses, dates, paths, filename encoding, summary counts and formulas. The PDF source stays unchanged. The CLI refuses ambiguous input and existing outputs rather than silently overwriting them.

**Scope:** this demonstrates a configured, text-layer document contract with document-local pagination and visible metadata. The client's eleven actual templates, image-only scans/OCR, handwritten/low-quality scans, security requirements and actual SharePoint synchronization have **not** been validated. `100%` describes this finite synthetic acceptance batch, not production accuracy. Types are placeholders until supplied by the client. PDF tables are retained visually; this milestone does not extract arbitrary table cells.

## A. Upwork proposal: Lee's Summit mailroom job

Target: [Automation needed to break out PDF into multiple files and extract data](https://www.upwork.com/freelance-jobs/apply/Automation-needed-break-out-PDF-into-multiple-files-and-extract-data_~022117380712874131495/).

The live page returned HTTP 403 during the original proposal preparation; the portability revision did not recheck the listing. This draft uses the job text supplied by the owner. Location, spend, proposal count and whether the listing remains open were not independently reverified; do not use them as sales claims.

**Paste-ready proposal**

Hi,

I built a small working demonstration of the workflow you described: one combined PDF becomes separate multi-page documents, an Excel breakdown of 11 document types, and files named from the person's name, status and document ID in a SharePoint-style folder tree.

The demonstration uses 15 synthetic documents across 24 pages. An independent parser checked every source/output document and page, with no missing pages or mismatched records in that test. I can share the source sample, extracted PDFs, Excel report and audit so you can inspect the result.

For your actual scanner output, I would first validate the boundaries and field rules against redacted examples of all 11 templates, including multi-page documents and repeats. Image-only scans need OCR; uncertain cases should go to review rather than being silently filed. The demonstration currently writes a local outbox. SharePoint upload or an approved synced folder would be configured and tested against your environment separately.

Could you share a redacted combined batch, the list of 11 types and statuses, and whether your scanner PDFs already contain searchable text? Also, should your SharePoint destination be a locally synced library or a direct integration?

I suggest a small paid validation milestone with agreed sample coverage and acceptance criteria before quoting full deployment. This keeps the first commitment measurable.

Thank you.

**Suggested commercial scope, for owner review:** propose $25/hour with a two-hour cap ($50) for initial template/OCR feasibility and a written implementation estimate. This is a proposed offer, not a market-rate claim or authorization to commit. Do not promise the complete 11-template OCR + SharePoint deployment for this amount.

## B. Zero-upfront-cost outreach drafts

No paid ads, purchased Connects, external APIs or cloud uploads were used for this PoC. Existing hardware and Python are assumed; public dependency installation is documented in README.md. Upwork proposals use Connects; the amount varies and purchasing Connects costs money. With zero starting cash, apply there only if the account already has enough eligible Connects or an applicable no-cost route. [Upwork's current Connects guidance](https://support.upwork.com/hc/en-us/articles/211062898-Understanding-and-using-Connects).

### r/forhire

**Title:** [For Hire] Remote PDF-to-Excel mailroom automation | $25/hour

I build focused PDF processing workflows: split combined batches into documents, extract configured fields, generate an Excel summary, and organize output filenames/folders.

My current demo processes 15 synthetic documents / 24 pages / 11 types. A separate parser audited all source and output pages and records with zero mismatches on this sample. The demo, workbook and audit are available to inspect. It supports a documented text-layer format; OCR and SharePoint deployment are scoped separately after reviewing your redacted samples.

Rate: $25/hour. A first paid feasibility milestone can be capped at two hours by agreement. Send a private message describing your document types, sample format and daily volume. I use AI-assisted development with independent automated checks and review; I do not claim untested production accuracy.

Before publishing, attach the actual synthetic proof pack or a published portfolio link and confirm account eligibility. The community requires an established active account, includes posting-frequency limits and directs job applications to private messages; the owner has not yet been checked for eligibility. [Community rules](https://www.reddit.com/r/forhire/comments/44aeko/rules_guidelines_read_before_posting/).

### r/slavelabour

**Title:** [OFFER] PDF batch splitting and Excel feasibility review — $25 fixed

For $25, I will review one redacted sample batch of up to 25 pages, identify whether its document boundaries/fields can be extracted reliably, and deliver a short feasibility report plus a scoped quote. This offer covers one hour of review, not complete OCR or SharePoint deployment.

I have a synthetic 15-document / 24-page demonstration with an independently checked Excel report and correctly separated PDFs. I use AI-assisted development and automated verification. If the source needs OCR or unsupported template work, I will identify that before proposing a larger paid milestone. Please follow the community's public `$bid` process before moving to private messages.

Check the live rules and account eligibility before posting. This is a paid, clearly bounded offer; do not advertise free custom work or send unsolicited bulk messages. [Full community rules](https://www.reddit.com/r/slavelabour/wiki/rules/).

### Direct reply to a relevant request

Hi — your combined-PDF-to-Excel workflow is close to a demo I have built. It separates multi-page documents, extracts configured identity/status fields, and creates a type/status summary with an independent audit. The attached example is synthetic, not a customer case. If you can share a redacted sample and expected results, I can propose a small paid validation step. OCR and your storage integration would be confirmed first.

Send this only where the recipient has invited contact; adapt it to the actual request.

### GitHub portfolio description

**Repository title:** Mailroom PDF splitter with independently audited Excel output

**README opening:** A local proof of concept for a mailroom batch: 15 synthetic documents, 24 pages, 11 configured types. Includes a repeatable CLI, generated test PDF, split output, a two-sheet workbook and an independent source-to-output audit. It is a template-driven text-layer demonstration; OCR and live SharePoint are not implemented. See the input contract and limitations before trying a new PDF.

Publish only synthetic artifacts and relevant project source. Exclude private state, chats, original client files, credentials, dependency environment folders and local machine paths. This portable edition uses public Python dependencies and includes the font redistribution license. Choose a project license before public publication; no repository was published in this session.

## Next paid milestone acceptance

Obtain approved redacted real scans covering each actual type and status. Agree page boundaries, required fields, repeated/ambiguous identifiers, unreadable-page handling and a review queue. Establish a withheld acceptance batch. Measure document-boundary precision/recall and field accuracy on that batch, with all pages accounted for and every uncertainty surfaced. Then validate the customer's SharePoint folder permissions, naming collisions, duplicate ingestion, retries and synchronization status. Confirm deployment dependencies and operating costs before a fixed implementation quote.

The backup [public-filings extraction job](https://www.upwork.com/freelance-jobs/apply/Python-Data-Scraping-Web-PDF-from-Public-Filings_~022114496291892509689/) is a separate opportunity supplied by the owner. This mailroom audit does not prove financial table extraction; do not reuse its accuracy claim for that job.
