# Branching And Roles

## Team Setup

- `codex`: builds the project structure, extracts assumptions, drafts the model logic, and prepares the integration baseline
- `cc`: can work on an alternative modeling implementation, polishing charts/tables, or a parallel writeup draft
- `libowen-integration`: used by Li Bowen's agent to reconcile differences and prepare the pre-submit candidate
- `main`: final reviewed version only

## Recommended Division Of Work

### Codex branch

- standardize repository structure
- extract the assignment requirements into a checklist
- document the cash flow model logic
- prepare scripts and templates for repeatable calculations
- produce the first full draft of the final submission package

### CC branch

- independently verify formulas and timeline assumptions
- produce alternative sensitivity scenarios
- improve narrative explanation and presentation quality
- identify weak assumptions or inconsistencies in the first draft

### Li Bowen integration

- compare `codex` and `cc`
- keep the stronger version when they differ
- resolve naming, formatting, and logic inconsistencies
- prepare the branch to merge into `main`

## Merge Order

1. `codex` establishes the baseline.
2. `cc` develops in parallel.
3. `libowen-integration` merges both branches and resolves conflicts.
4. Final review happens on `main`.

## Conflict Rule

If two branches change the same conclusion, keep the version with:

- clearer formula traceability
- assumptions explicitly documented
- better consistency with the assignment prompt
