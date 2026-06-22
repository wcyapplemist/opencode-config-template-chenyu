# pptx-subagent — User Stories

> These user stories follow the agile 3C principles (Card / Conversation /
> Confirmation). Each card is just a conversation starter — the real details
> are worked out through team discussion, not locked into this document.

---

## US-1: One-request deck generation

**As a** busy professional,
**I want** to describe my topic and page count in a single sentence and get back a finished presentation file,
**So that** I skip the hours of manual formatting, asset hunting, and layout tweaking.

**Acceptance Criteria:**
- Given a topic and a requested page count
- When the user submits the request in a single sentence
- Then the output is a file that opens directly in PowerPoint
- And every slide has a title and body content — no blank pages
- And the text and charts are individually editable, not flattened into pictures

**Discussion points:** How do we keep quality high for very long decks (20+ slides)? What's the default page count when the user doesn't specify one?

---

## US-2: Outline review checkpoint

**As a** content lead who needs to sign off on structure,
**I want** to see the full outline and approve it before any detailed content is written,
**So that** I can catch structural problems early instead of reworking finished slides.

**Acceptance Criteria:**
- Given an outline has been generated
- When the review checkpoint is reached
- Then generation pauses for approval or edits before continuing
- And the user may add/remove slides, reorder them, or change titles
- When running unattended with no reviewer available
- Then the agent judges the outline itself and proceeds

**Discussion points:** In unattended mode, what criteria does it use to decide the outline is good enough?

---

## US-3: Editable real charts

**As a** presenter who needs to show data,
**I want** the charts on my slides to be genuinely editable (not screenshots),
**So that** I can tweak the numbers anytime before the talk and the data looks credible.

**Acceptance Criteria:**
- Given a slide contains a chart
- When I double-click the chart in PowerPoint
- Then I can edit the values and the chart updates live
- And the chart's colors and styling match the rest of the deck — nothing looks out of place
- And data labels are visible (e.g., each bar shows its exact number on top)

**Discussion points:** Which chart types are supported? Who provides the data — does the user fill it in, or is it sourced automatically?

---

## US-4: Auto-sourced images

> **Status: Deferred — out of scope for the current phase (2026-06-19).** Stakeholder confirmed no auto-image-sourcing requirement at this stage. The auto-fetch code path (image_resolver / Pexels / Unsplash) has been removed; manual image embedding (`image_path`) remains. Re-open when auto-sourcing becomes a requirement.

**As a** presenter who wants a polished-looking deck,
**I want** relevant images to be found and placed automatically,
**So that** I don't have to search the web, download, and manually resize and position each picture.

**Acceptance Criteria:**
- Given the user describes the image they want
- When generation runs
- Then a suitable picture appears on the correct slide
- And its size and position are already set correctly — no manual dragging needed
- When no suitable image can be found
- Then the slide still generates without errors

**Discussion points:** Where do images come from (stock libraries vs. AI generation)? How is licensing handled?

---

## US-5: Speaker notes on every slide

**As a** someone who has to deliver the talk,
**I want** a spoken script attached to each slide,
**So that** I can rehearse beforehand and stay calm and fluent on stage.

**Acceptance Criteria:**
- Given a slide has been generated
- When the presenter opens Presenter View
- Then a full spoken script is visible (the audience doesn't see it)
- And the script is actual read-aloud sentences, not dry bullet summaries
- And each script includes the key point, a natural transition to the next slide, and at least one anticipated audience question

**Discussion points:** What language and tone should the scripts use? How long should each one be?

---

## US-6: Closing slide defaults to a professional sign-off

**As a** presenter who wants the talk to end well,
**I want** the final slide to default to a clean "Thank You" closing,
**So that** the presentation wraps up naturally rather than cutting off abruptly on a content slide.

**Acceptance Criteria:**
- Given the deck has 3+ slides and the user has not requested a custom closing
- When generation completes
- Then the last slide is automatically "Thank You"
- And the user never has to explicitly request a closing slide — it just happens
- And the closing slide's design matches the template, so it doesn't look tacked on

**Discussion points:** If the user wants a custom ending (e.g., "Q&A" or a tagline), how do they specify it?
