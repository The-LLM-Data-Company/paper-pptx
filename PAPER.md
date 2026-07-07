# paper-pptx Fork Ledger

Based on upstream tag `v1.0.2`, forked 2026-07-07, marker tag `paper-base`.

## Baseline Test Runs

Environment: CPython 3.12.13 via `uv`; test dependencies installed from upstream
`requirements-test.txt`.

- `pytest -q`: failed during collection with 39 errors, all caused by
  `pyparsing.warnings.PyparsingDeprecationWarning: 'delimitedList' deprecated - use 'DelimitedList'`.
  This is pre-existing environment drift from current `pyparsing` plus upstream's
  warnings-as-errors pytest configuration.
- `pytest -q -W ignore::pyparsing.warnings.PyparsingDeprecationWarning`: ran the
  suite and reported `2684 passed, 16 errors`; the remaining errors are current
  pytest fixture deprecation/error behavior around class-scoped fixtures defined
  as instance methods in `tests/opc/test_package.py` and
  `tests/opc/test_serialized.py`.
- `pytest -q -W ignore::pyparsing.warnings.PyparsingDeprecationWarning -W ignore::pytest.PytestRemovedIn10Warning`:
  `2700 passed`.
- `behave -q`: `54 features passed, 0 failed, 0 skipped`; `973 scenarios passed,
  0 failed, 0 skipped`; `2914 steps passed, 0 failed, 0 skipped`.
- `uv build`: built `dist/paper_pptx-0.1.0.tar.gz` and
  `dist/paper_pptx-0.1.0-py3-none-any.whl`. Setuptools emitted pre-existing
  license metadata deprecation warnings from upstream pyproject structure.
- Wheel smoke test:
  `uv run --isolated --no-project --with dist/*.whl python -c "import pptx; print(pptx.__paper_version__)"`
  printed `0.1.0`.
- Source distribution smoke test:
  `uv run --isolated --no-project --with dist/*.tar.gz python -c "import pptx; print(pptx.__paper_version__)"`
  printed `0.1.0`.

The CI workflow applies the two narrow pytest warning filters above when those
warning classes exist in the matrix environment, so current tooling can run the
upstream suite without modifying upstream source code.

## Publishing Safety

Publishing is intentionally disabled by default while this repository is
private. The release workflow targets the `pypi` environment and the publish
step is additionally guarded by `vars.PUBLISH_ENABLED == 'true'`. Configure
required reviewers on the `pypi` environment in GitHub before any release.

Do not push upstream `v*` tags to origin. Only the `paper-base` marker tag is
pushed during bootstrap.

## Sanctioned Deviations From Upstream Behavior

None.

## Upstream Merge Policy

Quarterly, run `git fetch upstream --tags`, identify whether a newer upstream
release tag exists, merge that release tag into `main`, and run both the pytest
and behave suites plus package smoke tests. Resolve conflicts using this file as
the map of intentional fork identity changes. Merge upstream releases; never
rebase `main`.
