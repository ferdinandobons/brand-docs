# Development

```bash
python3 -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt pytest
PYTHONPATH=scripts pytest -q        # docx / pptx / security / integration / smoke suites
```

> **Never commit real templates or company assets.** `brand-kit/` and `generated/`
> are intentionally git-ignored, and `tests/test_no_proprietary.py` fails the build
> if any Office binary is tracked outside `tests/fixtures` (or a vendored
> proprietary import sneaks in).

See [`CONTRIBUTING.md`](../CONTRIBUTING.md) and the frozen vocabulary in
[`CONVENTIONS.md`](../CONVENTIONS.md) before opening a PR.

## Release process

Release notes live in [`CHANGELOG.md`](../CHANGELOG.md). The shared engine,
profile schema, and QA gate are exercised by the CI suite
([`.github/workflows/ci.yml`](../.github/workflows/ci.yml)) on every push.
