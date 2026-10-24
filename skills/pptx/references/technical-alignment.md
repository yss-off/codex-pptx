# First-Touch Technical and Customer Alignment Decks

Use this pattern for first formal customer discussions, technical alignment, research-foundation presentations, delivery-boundary explanations, engineering demonstrations, and unresolved-requirement discussions.

## Mainline

Before polishing slides, write the meeting's one-sentence job. Test every slide against what it helps the audience believe, decide, or clarify.

A useful default sequence is:

1. Shared context and scope
2. Research or engineering foundation
3. Evidence already formed
4. Boundary between simulation, demonstration, and final delivery
5. Inputs or decisions that must be aligned next
6. Appendix for technical follow-up

If the meeting is not an acceptance review, avoid language that prematurely freezes final acceptance criteria. Use the user's terminology for technical positions, minimum metrics, delivery boundaries, and open inputs.

Ask before drafting when audience intent, delivery boundary, terminology, hardware role, live-versus-replay status, or meeting emphasis is materially uncertain.

## Slide jobs

| Slide role | Job | Necessity test |
|------------|-----|----------------|
| System context | Explain why the scoped work depends on the larger system | Does it support the deliverable without expanding scope? |
| Chain or block diagram | Split the deliverable into verifiable parts | Is it clear what each stage proves? |
| Demonstration or result | Present evidence rather than activity | Are numbers, plots, screenshots, and source mode tied to a claim? |
| Boundary | Prevent overclaiming | Are demonstration hardware and final hardware clearly separated? |
| Alignment request | Convert ambiguity into a concrete input | Will the audience know what answer is needed? |
| Appendix | Support technical follow-up | Is it reference material rather than duplicate body content? |

Body slides carry the persuasion line. Appendix slides support likely questions. Merge or remove appendix pages that only repeat the body.

## Terminology and proof layers

Create a terminology map and enforce it across the title, agenda, body, footer, diagrams, and appendix. Do not casually alternate terms for the same interface, migration activity, review status, or hardware role.

State proof layers explicitly:

- Simulation demonstrates modeled algorithm behavior under stated assumptions.
- SDR or USRP testing demonstrates recovery through an actual RF path under the tested conditions.
- Target-platform and final RF-board testing demonstrates closure on the delivery hardware.

Do not collapse these layers into one claim.

## Iteration with the user

When reviewing page by page, settle the page purpose, speaker message, keep/delete decision, and visual direction before editing that page or section. When a core term changes, scan the full deck for conflicts before finalizing.
