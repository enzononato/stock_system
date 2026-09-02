# AI Operating Instructions

## Project Understanding

Before modifying anything, thoroughly understand the relevant existing implementation.
Do not assume that a feature is missing before checking the code.

## Memory Protocol

At the beginning of every work session:

1. Read `instruções.md`.
2. Read `estado.md`.
3. Read `memoria.md`.
4. Read the latest relevant work plan from `notas/planos/`.
5. Inspect only the code relevant to the current task.

## Modification Rules

* Preserve existing functionality.
* Do not rewrite working systems without a clear reason.
* Do not remove functionality unless explicitly requested.
* Avoid unnecessary architectural changes.
* Respect existing technologies and project conventions.
* Keep changes focused on the requested task.
* Prefer small, reversible changes.
* Do not modify unrelated files.

## Validation

After significant modifications:

* Check for obvious errors.
* Validate affected functionality.
* Run the appropriate lightweight validation/build/test when practical.
* Never claim something works without validating it.

## Documentation Protocol

When a significant task is completed:

* Update `estado.md`.
* Update `memoria.md` if a permanent project decision changed.
* Create or update a work plan in `notas/planos/`.
* Create an audit in `notas/auditoria/`.

Do not create documentation for every tiny change.

## Planning Protocol

Before significant implementation:

* Create a checklist-based plan.
* Break the work into small steps.
* Mark completed items.
* Record blockers.
* Do not implement unrelated improvements.

## Audit Protocol

Audits must record:

* Scope
* Date
* Current status
* Findings
* Problems
* Risks
* Completed items
* Remaining items
* Recommended next actions

Keep audits concise.

## State Management

`estado.md` represents the CURRENT state of the project.
Do not turn it into a historical log.
Historical information belongs in audit files.

## Memory Management

`memoria.md` contains only information that future agents are likely to need.
Do not fill it with temporary details.

## Important Rule

When uncertain, inspect the relevant code instead of guessing.
Never invent project architecture, APIs, files, features, or implementation details.
