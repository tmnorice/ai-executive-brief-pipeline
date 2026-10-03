# AI Executive Brief + Meeting Prep Pipeline

A Google Apps Script and Gemini API workflow that gathers meeting context from Calendar, Gmail, Docs, and Sheets, classifies it with AI, and produces a sourced executive brief in Google Docs, with human review before anything is presented.

## Problem

Preparing an executive for a meeting usually means searching several places: the calendar invite, related email threads, project documents, and tracking sheets. The information an executive actually needs (decisions, risks, deadlines, and follow-ups) is buried inside all of it, and pulling it together by hand is slow and easy to get wrong.

## What the pipeline does

The workflow runs in stages from a control sheet. Each brief is a row in the `Briefing Queue`, and every stage processes only rows marked `Ready`.

1. **Pull context** from each enabled source into the `Source Log`, tagged with the Brief ID.
   - Calendar: the matching meeting event, attendees, and description
   - Gmail: threads matching the brief's search query
   - Google Docs and Sheets: the linked document text or sheet range
2. **Classify** each included source with the Gemini API into four fields: Decision Needed, Risk / Blocker, Deadline, and Follow-Up.
3. **Generate the brief**: a Google Doc with those four sections, a source reference for every item, and a human-review notice.
4. **Write back** the brief link, automation status, and run time to the queue, then set the row to `Complete`.

## Tools

- Google Sheets (control sheet)
- Google Apps Script
- Gemini API
- Google Calendar, Gmail, Google Docs, Google Drive

## Functions

| Function | Stage |
|---|---|
| `pullCalendarContext` | Pull: meeting event and attendees |
| `pullGmailContext` | Pull: matching email threads |
| `pullDriveContext` | Pull: linked Google Doc and Sheet range |
| `generateBriefSourcePacket` | Optional: compiles all raw sources into one review document |
| `classifyRemainingBriefSources` | Classify: Gemini extracts decisions, risks, deadlines, follow-ups |
| `generateExecutiveBrief` | Generate: builds the final brief for every Ready row |
| `testGeminiConnection`, `testGeminiClassification` | Diagnostics: confirm the API key and model respond |

Run order: pull stages → classify → generate. Each stage depends on the output of the one before it.

## Sheet structure

`Briefing Queue` expects these headers:

- Brief ID
- Meeting Date
- Meeting Time
- Meeting Title
- Executive / Owner
- Calendar Event ID
- Include Gmail
- Gmail Search Query
- Include Docs
- Docs Source URL
- Include Sheets
- Sheets Source URL
- Sheets Source Range
- Brief Status
- Brief Doc URL
- Automation Status
- Last Run
- Final Brief URL (created automatically on first run)

`Source Log` includes these headers:

- Brief ID
- Source Type
- Source Name / Subject
- Source URL / ID
- Date / Timestamp
- Owner / Sender
- Summary
- Decision Needed
- Risk / Blocker
- Deadline
- Follow-Up
- Included in Brief

## Setup

1. Create the two sheets with the headers above.
2. Open **Extensions → Apps Script** and add `Code.gs`.
3. Add the Gemini API key under **Project Settings → Script Properties** as `GEMINI_API_KEY`. The key is never stored in the code.
4. Run `testGeminiConnection` to confirm access.

## Design decisions

- **Human in the loop.** AI classifies; a person reviews the brief before it is distributed or acted on. Every brief ends with that notice.
- **Source traceability.** Every item in the brief lists the source it came from.
- **Status gating.** Only `Ready` rows are processed, and a finished brief sets its row to `Complete`, so completed work is not pulled or rebuilt again.
- **Duplicate prevention.** If a row already has a Final Brief URL, the brief is skipped rather than recreated.
- **Retry on temporary failures.** Classification retries API errors (429, 5xx) up to three times before logging the failure.

## Test cases

- `BRIEF-001`: the original build case. Four sources (Calendar, Gmail, Google Doc, Google Sheet) classified and compiled into a sourced brief.
- `BRIEF-002`: added after the classification and brief stages were refactored to read from the queue instead of a hard-coded Brief ID. A new Ready row with a sample Google Doc produced a second brief with no code changes, and BRIEF-001 was skipped because its brief already existed.

## Known limitations and next steps

- The pull stages update status but do not set rows to `Complete`; finished rows should be marked before re-running.
- Calendar matching uses the first event on the meeting date whose title matches.
- Next improvement: a single `runPipeline` function and a time-driven trigger so all stages run in order automatically.

## Notes

This repository contains a sanitized portfolio version of the workflow. It does not include private account credentials, client data, API keys, or internal business records. All test data is sample data.
