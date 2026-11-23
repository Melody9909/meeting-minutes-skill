# Meeting Content, Mapping, and Validation Rules

Apply these rules when extracting transcript content, validating structured JSON, and generating the editable Meeting Minutes DOCX. The confirmed field mapping in `template-schema.json` is authoritative.

## Source fidelity

- Treat the user-provided transcript as the factual source for meeting content.
- Preserve meaning, attribution, uncertainty, conditions, disagreement, and unresolved issues.
- Distinguish discussion, proposal, opinion, question, tentative plan, confirmed decision, and required action.
- Record a decision only when the transcript clearly confirms it.
- Record an Open Action only when the transcript supports a specific unfinished action or follow-up under a V4.1 category: a team commitment; work in progress with explicit remaining, continuing, or next work; a specific Client or Mentor requirement; a specific Client or Mentor improvement suggestion to consider; or a staged plan with an explicit follow-up action.
- Preserve whether an item is a suggestion, requirement, plan, commitment, or continuation of work in progress. Do not rewrite one nature or status as another.
- Do not infer facts or actions merely because they seem useful or likely.
- Do not create causal, conditional, evaluative-change, or other inferential relationships between statements unless the transcript supports that relationship explicitly.
- Remove verbal filler and repetition without changing meaning.
- Decide whether to retain arrangements between meeting participants by their relevance to the project, not by the participants' identities. Omit internal arrangements that do not concern the project's requirements, functionality, testing, progress, or matters the project team needs to address.
- Resolve names, roles, dates, and audiences only from user-provided material, except for schema-defined fixed values.

## Meeting role and attendance confirmation

- This Skill is specific to the Scholar Hub project. Use the fixed role mapping: **Basem Suleiman is the Client; Jinglin Sun is the Mentor.**
- Before extracting role-dependent content, identify the project team members and confirm which people actually attended this meeting from the user-provided material.
- A meeting name cannot establish attendance. In particular, `Regular Client Meeting` does not mean that Basem Suleiman or any Client attended.
- Do not use a professional title, hosting or arranging the meeting, testing the product, giving feedback, making requests, or guiding the team as attendance evidence or as a reason to change the fixed role mapping.
- Basem Suleiman's absence does not make Jinglin Sun the Client. A person's fixed role also does not establish attendance at every meeting.
- If attendance or project-team identity remains insufficient and would affect `questions_for_client` or `questions_for_mentor`, ask the user before classification instead of guessing.
- When no Client attended and there is no supported team-to-Client question, use `[]` for `questions_for_client`; Word must display `None` under the existing empty-section rule.

## Confirmed field mapping

### Fixed document values

- Keep `9900-F18A-ALMOND` unchanged.
- Keep `MEETING MINUTES` unchanged.
- Keep the `Meeting` heading unchanged; it is not a meeting-name placeholder.

### Meeting fields

- Produce every configured Meeting field as a string.
- Use the schema's fixed value `Liming Song` for `time_keeper`.
- Populate `meeting.time` only when the user directly provides the meeting time for manual entry. Do not extract or infer it from transcript content, transcript headers, speaker timestamps, file names, or file metadata. If the user does not provide it, use `[To be completed manually]`.
- For any other required Meeting field that cannot be confirmed, use `[Not stated in transcript]`.
- Do not treat unknown apologies as confirmation that nobody was absent.
- Do not invent materials, resources, links, a purpose, or a next meeting arrangement.
- Use `other_resources` for material that the transcript explicitly shows was provided, designated, or referenced in the meeting and that has independent value for follow-up work. This includes material explicitly shared for later reference or use, material the Client or Mentor explicitly directs participants to consult later, and a specifically identified resource whose supported purpose extends beyond the live meeting.
- Do not include a website link sent only to facilitate the live demonstration, slides, pages, or screen-shared content used only for the live presentation, ordinary mentions of resources with no supported follow-up purpose, or any resource whose follow-up value would have to be inferred. A link, presentation, or document may still qualify when the transcript supports its independent follow-up value; classify by supported purpose, not file or resource type.
- If the transcript explicitly requires participants to read or prepare material before the next meeting, assess that material for `please_bring_read` before `other_resources`. Do not repeat the same resource in both fields unless the transcript supports two distinct purposes and both are necessary to preserve meaning.

### Agenda

- Represent Agenda as a JSON array. Use `[]` when there are no supported agenda items.
- Each agenda entry contains `agenda_item`, `assignee`, and `notes` strings. Item numbers are generated sequentially from array order.
- `assignee` is only the person responsible for presenting or reporting that agenda item. It is not the person who first mentioned the topic and is not an action owner.
- If the presenter cannot be confirmed, use `[Not stated in transcript]`.
- `notes` may summarise material discussion and confirmed conclusions, but must preserve whether content was proposed, debated, decided, or unresolved.
- Preserve substantive project-related discussion, including questions, suggestions, requirements, and feedback from a confirmed Client or Mentor; their reasons, design thinking, specific methods, guidance, and limiting conditions; team explanations, discussion, general plans, decisions, and completed work; and unresolved matters with their related discussion.
- Do not omit an independent project-related matter solely because it appears less important. For long or compound statements, identify independent matters, requirements, conditions, intended audiences, methods, reasons, and statuses before summarising. Preserve key limitations and do not reduce distinct matters to a broader statement that changes their meaning.
- Reconcile project-relevant candidates against Agenda `notes`, including questions, suggestions, requirements, feedback, research, testing, reporting, and evaluation matters raised by the Client or Mentor and important project matters raised by other speakers. Each candidate must be covered or have a transcript-supported reason for exclusion. This is an internal check and is not an additional JSON field.
- Verbal repetition may be removed, but summarisation must not omit or merge away information with independent meaning.
- Keep the status of the same matter consistent across Agenda Notes by distinguishing suggestions, plans, commitments, work in progress, completed work, and unresolved matters. When the transcript explicitly shows a status change, preserve the actual change rather than mechanically assigning one status throughout.
- Preserve important discussions, questions, and answers in Agenda `notes`.
- When the team asks a confirmed Client or Mentor a project-relevant question, retain the necessary discussion under the relevant agenda item whether or not it was answered during the meeting.
- When a confirmed Client or Mentor asks the team a project-relevant question, retain both the question and the team's answer or related discussion under the relevant agenda item.
- For unresolved questions, accurately state their current status without inventing an answer or solution.
- Generate one Word row per agenda entry. If the array is empty, retain the Agenda table and enter `None` in its content row.

### Open Actions

- Represent Open Actions as a JSON array. Use `[]` when there are no supported action items.
- Each entry contains the action text in `action_discussed`. Keep separate actions in separate array entries and Word rows.
- Record an Open Action only when it is unfinished, has a transcript-supported action or follow-up with a clear object or bounded target, and falls into at least one of these categories:
  - a specific task that a team member explicitly commits to or accepts carrying out;
  - work already in progress for which the transcript identifies explicit remaining, continuing, or next work;
  - a specific requirement made by the Client or Mentor;
  - a specific improvement suggestion made by the Client or Mentor that calls for consideration; or
  - a staged plan that contains an explicit follow-up action.
- Ongoing or incomplete status alone is not enough. Do not include work in progress unless the original exchange identifies what is to be continued, completed, checked, revised, or otherwise followed up.
- Team commitment or acceptance is required when the action text represents a team commitment. It is not required merely to record a Client or Mentor requirement or suggestion, but the wording must preserve that source and nature and must not imply team commitment.
- Judge acceptance from the complete relevant dialogue, not a single sentence or fixed phrases such as “I will” or “We will”.
- When a Mentor or Client gives a clear, specific delivery requirement, a short team response such as “Okay” or “Yes” may support acceptance only if the response clearly refers to executing that task. A short response is not automatically a commitment; distinguish task acceptance from merely hearing, understanding, agreeing with a point, or continuing the discussion.
- When the team proposes a specific task, the Mentor or Client agrees or directs it to be done, and the team then confirms, assess the full exchange to determine whether execution consensus was formed.
- A Mentor or Client suggestion, requirement, request, or command does not by itself establish a team commitment. It may nevertheless qualify in its own right as a V4.1 Open Action when it is specific and has a clear follow-up. Phrase it so its original nature remains explicit; for example, use `Consider adding...` for an improvement suggestion rather than rewriting it as `Add...`.
- Statements such as “we will consider it” or “we will investigate feasibility” qualify only when the transcript identifies a specific object or bounded target for that consideration or investigation. Preserve the limited commitment to consider or investigate; do not rewrite it as a commitment to implement.
- For compound team statements, split independent clauses and matters before classifying actions. When one statement contains completed work, work in progress, and a future commitment, preserve and evaluate each status separately. Do not replace an explicit commitment with a broad summary that loses its commitment status.
- For each candidate, confirm from the original statement that the action is explicit, the object or bounded target of the action is explicit, and the task can be understood without adding task content that the speaker did not provide. Specificity concerns what work is committed and what it applies to; it does not require the implementation method or solution to have been decided. A broad future improvement direction remains a plan when the original wording does not identify a sufficiently bounded object or target for the improvement, even if it uses “we will”, names a future Sprint, or gives a time.
- Do not make a broad plan appear specific by rewriting it, resolving ambiguity through context, or adding a deliverable, scope, target, or other task content not stated by the speaker. Specificity does not mechanically require an implementation method, quantity, deadline, or acceptance criterion; it requires a sufficiently clear original action and object or bounded target.
- Keep broad suggestions, proposals, exploratory statements, and plans without explicit follow-up in Agenda `notes`, accurately preserving whether they were suggested, proposed, under consideration, planned, or unresolved. A specific Client or Mentor suggestion or staged plan with explicit follow-up may also qualify for Open Actions, but its original nature must be retained.
- Do not infer or create an action from context, usefulness, common practice, or expected next steps.
- Build a separate internal Open Actions candidate set. Reverse-scan relevant exchanges for all V4.1 categories, considering the speaker and role, the item's original nature, the action or follow-up, its object or bounded scope, unfinished status, and task-directed acceptance evidence when the item is represented as a team commitment. Account for every candidate by mapping it to a final Open Action or recording a transcript-supported reason for exclusion. This internal check is not an additional JSON field.
- Record action content only. Do not include an executor, owner, or responsibility label, even when the transcript names one.
- Set `deadline` to the empty string `""` for every action, even when the transcript states a date. The field is reserved for later manual completion.
- Item numbers are generated sequentially from array order.
- If the array is empty, retain the Open Actions table and enter `None` in its content row; keep Deadline blank.

### Agenda and Open Actions classification

- Keep the discussion background, problems, reasons, suggestions, requirements, feedback, decisions, and current status in Agenda `notes`. Put a qualifying specific follow-up in `open_actions` while preserving whether it is a suggestion, requirement, plan, commitment, or continuation of work in progress.
- The same matter may appear in Agenda `notes` and `open_actions` when the Agenda entry provides useful discussion or status context and the Open Action records the concrete follow-up. Do not treat this as an error, but avoid identical wording or repetition that adds no record value.
- Keep completed work, broad plans, broad discussion, and exploratory ideas without explicit follow-up in Agenda `notes`; do not force them into Open Actions.

### Questions

- Represent `questions_for_mentor` and `questions_for_client` as separate arrays of strings.
- In `questions_for_client`, record questions the team asked the Client during the meeting only when they have independent project-record value, whether or not the Client answered or confirmed them. This includes questions concerning product requirements, functional design, business rules, project decisions, Client feedback, or business matters requiring confirmation.
- In `questions_for_mentor`, record questions the team asked the Mentor only when they have independent project-record value, whether or not the Mentor answered or confirmed them. Relevant subjects include project progress, deliverables, files and reports, testing, acceptance, project arrangements, Mentor feedback, and other matters whose record assists project work; these examples are not a closed list.
- A question has independent project-record value when, after the meeting, its subject and response, confirmation, or unresolved status still help explain project requirements, advance project work, or support later reference. Do not exclude a clarification question merely because it is brief; assess its underlying matter in context.
- Exclude questions used only to manage the immediate interaction, such as requests for repetition, sound checks, or temporary demonstration-operation confirmations, when they have no independent project-record value.
- Do not rewrite a statement, suggestion, conditional offer, or Mentor requirement as a team question. Keep Mentor-initiated project-delivery requirements in Agenda `notes`; add an Open Action only when the Open Actions conditions are also satisfied.
- Classify a question by the person or role it was addressed to, not by whether it was resolved.
- Keep questions raised by the Client to the team in Agenda `notes`; do not place them in `questions_for_client`.
- A team question may appear both in Agenda `notes` and in the applicable Questions array because the sections serve different purposes.
- Add a question only when the transcript supports both the question and its intended respondent. Apply the fixed role mapping, but do not guess attendance or whom an ambiguous question addressed; when that uncertainty affects classification, ask the user first.
- Use `[]` for an empty Questions array. In Word, preserve the corresponding section and display `None`.

## Uncertain or conflicting content

- Do not guess inaudible, ambiguous, contradictory, or missing content.
- Use `[Unclear in transcript]` for a value that is present but cannot be resolved accurately.
- Use `[Not stated in transcript]` for a required Meeting or Agenda string that has no supported value.
- Use an empty array, not a missing-value marker, when Agenda, Open Actions, or a Questions section has no supported entries.
- If conflicting statements remain unresolved, record the conflict without choosing a side.

## JSON validation

Before bidirectional verification, save the complete first-extraction JSON to a distinct file. Keep that file unchanged as the comparison baseline; corrections belong in a separate working or final JSON file.

Perform both verification directions against the complete original transcript before Schema validation:

- **Source-to-output completeness:** after first extraction, reread the complete original transcript from beginning to end. For every segment, check Agenda Notes, Open Actions, Questions for Mentor, and Questions for Client. Find omitted independent project matters, actionable follow-up, record-worthy team questions and answers, key conditions, and classification errors. Reviewing only the candidate list or existing output is not sufficient.
- **Output-to-source fidelity:** check every Agenda Note, Open Action, Mentor question, and Client question against its supporting original exchange. Confirm transcript support, attribution, meaning, original nature, status, and classification, and confirm that no unsupported information was added.
- When either direction finds an omission, unsupported statement, altered meaning or status, or classification error, correct the working JSON first and repeat the affected checks against the transcript. Do not merely describe a known defect in a test report.
- Preserve the first-extraction JSON unchanged. Only after corrections and both verification directions pass may the corrected JSON be validated and saved as the final JSON used for DOCX generation.

Before DOCX generation, verify the corrected final JSON structurally:

- the root data keys are `meeting`, `agenda`, `open_actions`, `questions_for_mentor`, and `questions_for_client`;
- `meeting` contains exactly the configured Meeting keys and all values are strings;
- `agenda`, `open_actions`, `questions_for_mentor`, and `questions_for_client` are arrays;
- each array entry uses only its configured keys and types;
- required strings are not empty after applying fixed and missing values;
- every Open Actions deadline is `""`;
- no action text contains an owner added from the transcript;
- empty content sections use `[]`; and
- no unconfigured field or section has been added.

If validation fails, patch only the invalid fields in the corrected working JSON and revalidate them against the transcript and schema. Do not modify the saved first-extraction JSON.

## DOCX generation and verification

- Generate the DOCX only from the corrected final JSON after bidirectional verification and Schema validation have passed.
- Keep the Tutor DOCX in `assets/` unchanged. Use the confirmed schema as the routine structural mapping and the Tutor DOCX as the reference for section and field architecture.
- Preserve section order, field labels, table columns, and the separate Mentor and Client question sections.
- Add or remove Agenda and Open Actions body rows to match the validated arrays.
- When a content array is empty, retain its Word section or table and show `None` as specified above.
- The final DOCX must remain editable and open without corruption or repair warnings.
- Verify that rendered document content matches the validated JSON and that no value was dropped, duplicated, or moved to the wrong field.
- Exact fonts, colours, borders, column widths, and pagination do not need to match the original Tutor template.
- Do not require PDF or PNG conversion for routine generation or acceptance. Use conversion only when necessary to diagnose a document problem or when the user requests it.

## Final content checks

Before delivery, confirm that:

- every substantive statement is traceable to the transcript or is a schema-defined fixed value;
- `meeting.time` uses a time directly provided by the user for manual entry, or `[To be completed manually]` when the user did not provide one; it has not been extracted or inferred from the transcript or file information;
- discussions, decisions, questions, and actions have not been conflated;
- the complete first-extraction JSON was saved separately before bidirectional verification and remains unchanged;
- the complete original transcript was reread from beginning to end for source-to-output completeness across Agenda Notes, Open Actions, Questions for Mentor, and Questions for Client, rather than checking only existing candidates;
- every output item in those four sections was checked back against its supporting original exchange for source support, attribution, meaning, original nature, status, and classification;
- every Open Action is unfinished, contains a transcript-supported specific action or follow-up and clear object or bounded target, and qualifies as a team commitment, work in progress with explicit remaining or continuing work, a specific Client or Mentor requirement, a specific Client or Mentor improvement suggestion to consider, or a staged plan with explicit follow-up;
- no ongoing work was included merely because it was incomplete, and every included in-progress item states transcript-supported remaining, continuing, or next work;
- a unilateral Client or Mentor suggestion, requirement, request, or command is not represented as a team commitment; any short response used as acceptance is clearly tied by context to execution of the specific task;
- every suggestion, requirement, plan, commitment, and in-progress item retains its original nature, and no suggestion was rewritten as a commitment, no plan as a decision, and no in-progress work as completed;
- compound team statements have been split into independent clauses and matters, and completed, in-progress, and future-commitment statuses within the same statement have been evaluated separately without replacing an explicit commitment with a broad summary;
- separate internal Agenda Notes, Open Actions, Questions for Mentor, and Questions for Client candidates have been checked, and a reverse scan of relevant exchanges has accounted for every V4.1 Open Action candidate through a final Open Action or a transcript-supported exclusion reason;
- every Open Action's original statement supplies a sufficiently clear action and object without unsupported task content, and no broad plan has been converted into a specific task through rewriting, inference, future-tense wording, a future Sprint, or timing alone;
- Agenda Notes preserve important discussions, team-to-Client questions, Client-to-team questions, answers, and the current status of unresolved questions;
- Agenda Notes preserve substantive project-related questions, answers, suggestions, requirements, feedback, reasons, design thinking, methods, guidance, conditions, team explanations, general plans, decisions, completed work, current functionality, and unresolved matters without losing distinct meaning through over-summarisation or subjective importance filtering;
- long or compound statements have been checked for independent matters, requirements, conditions, intended audiences, methods, reasons, and statuses, and key limitations have not been dropped;
- project-relevant candidates have been reconciled against Agenda Notes, with each candidate covered or excluded for a transcript-supported reason;
- the same matter has a consistent status across Notes unless the transcript explicitly records a real status change;
- qualifying follow-up is recorded in Open Actions, while Agenda Notes retain useful discussion background, problems, reasons, feedback, decisions, and status; meaningful overlap is allowed, but identical or otherwise valueless repetition is removed;
- every omission, unsupported statement, meaning or status change, and classification error found during bidirectional verification was corrected in the working JSON before Schema validation and DOCX generation, rather than being left only in a test report;
- each `other_resources` entry has transcript evidence that the material was explicitly provided, designated, or referenced and has independent follow-up value; live-demo-only material, ordinary unsupported resource mentions, and inferred resources are excluded;
- material explicitly required for reading or preparation before the next meeting has been assessed for `please_bring_read`, and the same resource is not repeated across `please_bring_read` and `other_resources` without a distinct transcript-supported purpose;
- no unsupported causal, conditional, evaluative-change, or other inferential relationship has been introduced;
- the fixed Scholar Hub mapping was applied (Basem Suleiman as Client and Jinglin Sun as Mentor), while project team membership and actual attendance were confirmed separately from supplied material; the meeting name, titles, and participant behaviour were not used as attendance evidence or to change those fixed roles;
- every question in `questions_for_client` was asked by the team to the confirmed Client, has independent project-record value, and is supported as an actual question rather than a rewritten statement, suggestion, or conditional offer, regardless of whether it was answered or resolved;
- no team-to-Client question with independent project-record value was excluded merely because it was answered, brief, or phrased as a clarification;
- every question in `questions_for_mentor` was asked by the team to the confirmed Mentor, has independent project-record value, and remains useful after the meeting for understanding requirements, advancing work, or later reference, regardless of whether it was answered or resolved;
- temporary interaction-management questions without independent project-record value were excluded from both Questions sections;
- Client-to-team questions have been retained in Agenda Notes and not placed in `questions_for_client`;
- duplication of a team question between Agenda `notes` and its applicable Questions array is not treated as an error;
- Agenda presenters are supported and are not action owners by default;
- Open Actions contain no owners and all Deadline cells are blank;
- dynamic row counts match the validated arrays;
- empty arrays render as `None` in the correct Word section;
- `9900-F18A-ALMOND`, `MEETING MINUTES`, and `Meeting` remain unchanged;
- the document contains no extra field or section; and
- the DOCX is editable, structurally complete, and not damaged.
