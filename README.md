# paper-pptx

`paper-pptx` is Paper Instruments' hard fork of
[`python-pptx`](https://github.com/scanny/python-pptx), based on the upstream
`v1.0.2` release.

This fork is intended to remain a strict superset of upstream behavior. Existing
code such as `from pptx import Presentation` must continue to work unchanged.

## Naming

There are four names to keep distinct:

- GitHub repository: `paper-pptx`
- PyPI distribution: `paper-pptx`
- Python import package: `pptx`
- Fork sentinel: `pptx.__paper_version__`

Built wheel files are named `paper_pptx-*`, while the import remains `pptx`.
That mismatch is intentional and follows the normal distribution/import split
used by packages such as Pillow/PIL. Do not rename `src/pptx` to `src/paper_pptx`.

## Installation

This repository is private and publication to PyPI is intentionally gated. For
now, install from Git:

```bash
pip install "paper-pptx @ git+https://github.com/The-LLM-Data-Company/paper-pptx.git@main"
```

Verify the fork sentinel:

```bash
python -c "import pptx; print(pptx.__paper_version__)"
```

See `PAPER.md` for fork lineage, baseline test results, and merge policy.
