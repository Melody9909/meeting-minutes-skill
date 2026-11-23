---
name: meeting-minutes
description: Create English meeting minutes for the personal Scholar Hub project from an English meeting transcript, using the fixed school Tutor template and fixed Scholar Hub role mapping. Use when the user asks to extract discussions, decisions, and action items into that template.
---

# Meeting Minutes

Create accurate English Meeting Minutes from an English meeting transcript.

## Required inputs

- An English meeting transcript supplied by the user.
- The school's original Tutor template stored in `assets/`.

If either input is unavailable or unreadable, stop and ask the user to provide it. Do not invent or substitute a template.

## Runtime requirements

- Python 3.10 or later.
- Install the verified Python dependency from `requirements.txt`.

## Required V4.1 execution path

For every meeting-minutes generation task, preserve all confirmed field meanings, fixed values, dynamic-row behavior, and empty-value rules, and use this complete execution path:

1. Read the meeting transcript.
2. Read `references/meeting-rules.md` and extract only content supported by the transcript.
   - This Skill is specific to Scholar Hub: Basem Suleiman is the Client and Jinglin Sun is the Mentor. Before extraction, identify the project team members and confirm which people actually attended this meeting from the supplied material. A meeting name, professional title, meeting-host behaviour, feedback, or guidance is not attendance evidence. Do not assume that either fixed role attended, and never reclassify Jinglin Sun as the Client because Basem Suleiman was absent. If unresolved attendance or team-member identity affects Questions for Client or Questions for Mentor, ask the user before classifying those questions.
   - For Questions for Client and Questions for Mentor, retain team questions only when they have independent project-record value, including relevant questions answered during the meeting. For Mentor questions, this includes matters such as project progress, deliverables, files and reports, testing, acceptance, project arrangements, and Mentor feedback, without limiting the category to those examples. Exclude temporary interaction-management questions that have no value after the meeting, and do not rewrite a Mentor statement, suggestion, or requirement as a team question.
   - For long or compound statements, identify each independent clause and project matter, including its requirement, condition, intended audience, method, and status, before summarising. If one team statement mixes completed work, work in progress, and a future commitment, classify each clause separately. Preserve transcript-supported reasons, design thinking, guidance, and limiting conditions; do not omit an independent project matter solely because it appears less important or replace an explicit commitment with a broader summary that loses its commitment status.
   - Build separate internal candidate sets for Agenda Notes, Open Actions, Questions for Mentor, and Questions for Client without adding JSON fields. Reconcile project-relevant candidates—including questions, suggestions, requirements, feedback, research, testing, reporting, and evaluation matters raised by the Client or Mentor, as well as important project matters raised by other speakers—against the appropriate output section. Each candidate must be covered or have a transcript-supported reason for exclusion.
   - For every Open Action candidate, verify the complete relevant exchange, the specific executable or follow-up matter and its object or bounded target, its unfinished status, and its original nature. Open Actions may include: a specific task the team explicitly committed to; work already in progress with explicit remaining or continuing work; a specific Client or Mentor requirement; a specific Client or Mentor improvement suggestion that calls for consideration; or a staged plan with an explicit follow-up action. Team acceptance is required only when the output represents a team commitment; it is not required merely to record a Client or Mentor requirement or suggestion. Do not turn acknowledgment into commitment, a suggestion into a requirement or commitment, a plan into a decision, or work in progress into completed work. Do not include ongoing work merely because it is unfinished: the transcript must identify clear remaining, continuing, or next work. Do not infer an action from a broad discussion, general direction, expected next step, or usefulness. Reverse-scan candidate exchanges rather than keywords alone; every candidate must map to a final Open Action or have a transcript-supported reason for exclusion.
   - Extract `please_bring_read` and `other_resources` separately. Use `other_resources` only for material explicitly provided, designated, or referenced in the meeting that has independent value for follow-up work; exclude material used only to facilitate the live demonstration and ordinary resource mentions with no supported follow-up value. If participants are explicitly required to read or prepare material before the next meeting, assess it for `please_bring_read` first and avoid meaningless duplication between the fields.
3. Build structured meeting JSON conforming to `references/template-schema.json`.
4. Save the complete first-extraction JSON to a distinct file before bidirectional verification. Do not overwrite this file with the corrected or final JSON; it is the baseline for later comparison.
5. Perform bidirectional verification against the complete original transcript before Schema validation or Word generation. Candidate-list review alone is not sufficient.
   - **Source-to-output completeness:** reread the complete original transcript from beginning to end after the first extraction. Check every segment against Agenda Notes, Open Actions, Questions for Mentor, and Questions for Client, and confirm that no independent project matter, actionable follow-up, record-worthy team question, answer, or key limiting condition has been omitted or misclassified. Repeated expressions may be combined, but every candidate must be covered or have a transcript-supported reason for exclusion.
   - **Output-to-source fidelity:** check every Agenda Note, Open Action, Mentor question, and Client question against its supporting original exchange. Confirm that it has transcript support, adds no unsupported information, preserves meaning and attribution, retains the original nature and status, and is in the correct section.
   - **Status consistency:** check the same matter across all Notes and preserve whether it was a suggestion, plan, commitment, work in progress, completed work, or unresolved. If the transcript explicitly shows a status change, preserve the actual change rather than forcing one status throughout.
   - **Classification and useful overlap:** confirm that each Open Action qualifies under one of the supported V4.1 categories and preserves whether it is a suggestion, requirement, plan, commitment, or continuation of work in progress. Exclude completed work, broad discussion without explicit follow-up, and ongoing work without clear remaining, continuing, or next work. Confirm that no broad matter was made artificially specific by paraphrasing or adding unsupported task content. The same matter may appear in Agenda Notes and Open Actions when Agenda Notes preserve useful background, reasons, status, feedback, decision context, or discussion and Open Actions preserve the concrete follow-up. Avoid identical or otherwise meaningless repetition.
   - **Meeting resources:** confirm that every `other_resources` entry is supported by the transcript as material with independent follow-up value, that live-demo-only material and unsupported resource mentions are excluded, and that material required for preparation before the next meeting has been assessed for `please_bring_read` without unnecessary duplication.
6. If either verification direction finds an omission, unsupported statement, meaning or status change, or classification error, correct the structured JSON first. Preserve the first-extraction JSON unchanged. Repeat the affected checks against the original transcript until they pass.
7. Validate the corrected JSON structure, field types, required values, fixed values, and empty values against the Schema, then save the corrected, source-verified JSON to a separate final file path.
8. Generate the editable DOCX exclusively through `scripts/generate_minutes.py`.
   - Resolve the installed Skill directory to an absolute path; do not depend on the current working directory.
   - Use actual absolute paths for the validated meeting JSON and output DOCX.
   - Ensure the generator resolves or receives valid absolute paths for `references/template-schema.json`, `references/meeting-rules.md`, and the Tutor template under `assets/`.
   - Do not write a meeting-specific PowerShell script, use Microsoft Word COM, or substitute another JSON-to-DOCX generator.
9. Reopen and inspect the generated DOCX for editability, structural completeness, and consistency with the validated final JSON.

The following failure rules are mandatory:

- If JSON Schema validation fails, correct only the invalid fields and validate again. Do not re-extract unaffected transcript content.
- If bidirectional verification fails, correct the affected content in the working JSON, verify it against the complete original transcript again, and then validate the corrected JSON. Do not merely record the defect in a test report, do not alter the saved first-extraction JSON, and do not generate the final DOCX from a JSON file with a known defect.
- If DOCX generation fails, retain the validated JSON and retry only `scripts/generate_minutes.py`, or report the generation error. Do not re-extract the meeting transcript.
- Do not require PDF or PNG conversion, and do not use Microsoft Word COM.

## Workflow

1. Read the complete transcript without altering the source file.
2. Read [references/meeting-rules.md](references/meeting-rules.md) before extracting or drafting content.
3. Treat the Tutor Word template in `assets/` as an immutable source. Create the completed document from a working copy; never edit or overwrite the original template.
4. Apply the fixed Scholar Hub role mapping and confirm actual attendance and project team membership using `references/meeting-rules.md`; ask the user when unresolved attendance or team-member identity would affect question classification. Then split compound statements into independent clauses and project matters, and build separate internal Agenda Notes, Open Actions, Questions for Mentor, and Questions for Client candidates. Preserve substantive project discussion without relying on a subjective importance threshold. Check team-to-Client and team-to-Mentor questions for independent project-record value, including answered questions; check all project-relevant requirements, feedback, reasons, methods, guidance, and conditions; and check every V4.1 Open Action category using the complete relevant exchange. Classify suggestions, requirements, plans, commitments, work in progress, completed work, and unresolved matters separately when they occur in the same statement. Treat work in progress as an Open Action only when the transcript identifies explicit remaining, continuing, or next work. Extract Meeting-resource candidates separately under the `please_bring_read` and `other_resources` rules.
5. Populate the Tutor template without adding, removing, renaming, or reordering its fields or sections.
6. Save the first-extraction JSON as an immutable comparison baseline. Then conduct the source-to-output and output-to-source checks against the complete original transcript for Agenda Notes, Open Actions, Questions for Mentor, and Questions for Client. Verify every Open Action's category, original nature, action or follow-up, object, specificity, and unfinished status; verify team acceptance only when representing a team commitment; confirm that ongoing work has explicit remaining, continuing, or next work; confirm that broad matters were not rewritten into specific tasks; allow useful Agenda/Open Actions overlap while removing meaningless repetition; verify that `other_resources` has transcript support for independent follow-up value and does not duplicate `please_bring_read` without purpose; and reject unsupported inference or contradictory status. Correct the working JSON before validation whenever a check fails.
7. Validate and save the corrected final JSON separately, generate the DOCX from that file, and return only the completed English Meeting Minutes in the format required by the Tutor template.

## Tutor template contract

- Preserve `9900-F18A-ALMOND` exactly as written. Never replace or remove it.
- Preserve the `Meeting` heading exactly as written. Do not replace it with a meeting name.
- In the Agenda table, use `Assignee` only for the person responsible for presenting that agenda item. Do not use it for an action owner.
- Preserve the existing numbered-list formatting under `Questions for Mentor` and `Questions for Client`. If the transcript contains no relevant question for one of these sections, enter `None` as its numbered-list content.
- Keep the Open Actions fields unchanged. Record only the action content; do not add or include an execution owner. Always leave `Deadline` empty for later manual completion.

## Boundaries

- Preserve the meaning of the meeting rather than rewriting it into unsupported conclusions.
- Never add facts, decisions, owners, deadlines, attendees, or context that the transcript does not support.
- Do not design a new Meeting Minutes format or infer missing template fields.
- Do not add, delete, rename, or reorder any template heading, field, table column, or section.
- Do not generate a Chinese task list, bilingual summary, commentary, or extra sections outside the Tutor template.
- Do not create an example Meeting Minutes document unless the user explicitly requests one and provides suitable source material.

