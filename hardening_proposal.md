# paper-pptx Practical Hardening Proposal

## Objective

Harden the operations that professionals rely on when editing an existing presentation: replace content, modify structure, compose decks, inspect the result, and save it without damaging unrelated content.

This proposal is intentionally narrow. It includes only failures that can produce a wrong deck, an incomplete rollback, a misleading result, or an invalid package during practical use. It does not attempt to reject every theoretically malformed XML shape or defend every unsupported manipulation of private internals.

## Release standard

For the operations covered here, `paper-pptx` should meet three requirements:

1. Validate the complete operation before changing the presentation.
2. On refusal, restore the original live object graph as well as the serialized package.
3. On success, preserve unrelated presentation content and report the change accurately.

## Proposed work

### 1. Make failed edits genuinely atomic

Some multi-part operations can fail after changing XML elements, relationships, or binary part data. Restoring only serialized bytes is insufficient because callers may continue using existing `Presentation`, slide, shape, chart, or table objects after the refusal.

Implement one transaction boundary for public mutators that:

- captures every XML tree, relationship collection, and mutable binary part involved in the operation;
- performs candidate-package validation without mutating the candidate or the live presentation;
- restores the original live objects and cached proxies when validation or serialization fails; and
- supports nested operations without taking a stale snapshot or partially committing an inner operation.

This is the highest-priority item because atomic refusal is part of the package's central contract. A banker or lawyer should be able to catch a typed refusal and continue working with the same in-memory presentation safely.

**Acceptance criteria**

- Injected failures at each mutation and save stage leave both the in-memory presentation and the file on disk unchanged.
- Existing slide, shape, table, picture, and chart proxies either remain valid or refuse as stale; they never silently edit detached XML.
- Custom XML and binary parts are restored along with standard PowerPoint parts.

### 2. Protect relationship and shared-part integrity

PowerPoint objects are connected by package relationships. An edit that validates only the visible XML can leave a dangling reference, preserve an unintended alias, or overwrite data shared by another object.

Before picture replacement, chart-data replacement, slide or layout deletion, and related package surgery:

- require relationship IDs in XML and relationship collections to agree;
- verify that the target part has the expected type and ownership;
- detect additional relationships that alias the same slide, layout, image, chart, or embedded workbook;
- refuse in-place replacement when an embedded workbook or other mutable part is shared by another object; and
- remove relationships only after proving that no surviving XML references them.

This prevents decks that open successfully while a chart points at inconsistent workbook data, an image reference is dangling, or supposedly deleted content remains hidden in the package.

**Acceptance criteria**

- Saving after each supported edit produces no dangling internal relationship.
- Replacing one chart's data cannot alter another chart through a shared embedded workbook.
- Deleting a slide or layout either removes its complete owned subgraph or refuses before mutation.

### 3. Reject stale editing targets

Long-running workflows commonly retain references to slides and shapes while other operations restructure the deck. A proxy whose XML node has been replaced or detached must not appear to succeed while changing content that will not be saved.

Add consistent ownership checks to public mutators for slides, layouts, shapes, text frames, tables, pictures, charts, and header/footer objects. The check must prove that the proxy still identifies the active object in the current presentation, not merely that it once belonged to the same package.

**Acceptance criteria**

- Mutating a detached or superseded proxy raises a typed refusal before changing any state.
- Normal proxies remain usable after unrelated edits.
- Composition, rebind, clone, and deletion operations explicitly invalidate only the proxies they replace.

### 4. Make inspection and diff results trustworthy

Inspection and verification are decision inputs for automated workflows. Missing inherited formatting, notes, geometry changes, or source-package changes can cause a caller to approve the wrong deck.

Harden the existing inspection and diff APIs so that they:

- resolve effective text and shape properties from the correct placeholder, layout, master, and theme chain;
- report provenance without silently choosing among ambiguous sources;
- include supported notes, tables, charts, images, fields, and inherited geometry consistently;
- compare the actual supplied packages rather than accidentally reusing live or cached state; and
- refuse malformed or ambiguous structures when a deterministic answer is unavailable.

The goal is not broader reporting. It is that every value already promised by the APIs is complete and dependable enough to drive an automated review step.

**Acceptance criteria**

- Golden fixtures cover inherited formatting, notes, tables, chart data, images, fields, slide movement, layout rebinding, and geometry changes.
- Repeating inspection or diffing produces deterministic output.
- Diffing two paths, streams, or live presentations compares those exact inputs and restores stream positions afterward.

### 5. Preserve identity and ownership during deck surgery

Slide cloning, cross-deck import, layout rebinding, and structural deletion touch many related parts. Incorrect identity remapping or incomplete dependency analysis can create duplicate creation IDs, shared mutable content, or references back to the source package.

For these operations:

- inventory the full relationship subgraph before mutation;
- assign unique document-wide identities to every independently copied object and part;
- preserve shared immutable media only when the selected policy permits it;
- ensure copied mutable parts are independent;
- validate slide, layout, master, notes, comments, charts, and embedded-workbook ownership; and
- refuse unsupported reconciliation before adding any destination parts.

**Acceptance criteria**

- Imported or cloned slides have no relationship to the source package after the operation.
- Independently copied parts never receive duplicate document-wide creation IDs.
- Editing a cloned or imported chart, image, note, or table does not change its source counterpart.
- A failed compose, clone, rebind, or delete operation leaves the destination byte-identical and its existing proxies usable.

## Verification required for release

The implementation should be delivered as small commits organized by the five workstreams above. Each commit should include a regression test that fails against `v0.1.1` and passes with the fix.

Before release:

- run the complete Paper contract suite and inherited `python-pptx` suite;
- build the wheel and source distribution from a clean checkout;
- install the wheel into an empty environment and run the import doctor;
- exercise representative inspect, replace, clone, compose, diff, refuse, and save workflows;
- open every resulting fixture with both PowerPoint-compatible package validation and the LibreOffice load smoke; and
- confirm refusal atomicity for in-memory objects, output paths, and file-like streams.

## Explicitly out of scope

- Exhaustive validation of every malformed OOXML permutation.
- New authoring features or new public APIs.
- SmartArt, animation, transition, rendering, or slide-size migration work.
- Refactors whose only benefit is cleaner internal architecture.
- Additional refusal cases that require deliberate mutation of private attributes and cannot affect supported public workflows.
- Documentation, CI, packaging, or install changes unless required to test and ship one of the correctness fixes above.

These exclusions are deliberate. The target is a smaller, reviewable hardening release that materially reduces the chance of silent damage in real presentation workflows.
