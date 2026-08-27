# Developer guide

## Local checks

```bash
python -m pip install -e .
python -m unittest discover -s tests -v
python -m compileall -q src tests packaging
```

Run an end-to-end refresh against a source root:

```bash
python -m wfmhub --home . setup --source-root /path/to/source/root --non-interactive
python -m wfmhub --home . refresh --start 2026-08-01 --end 2026-08-31
```

Local config, extracts, DuckDB, reports, logs and portable build products are
gitignored.

## Add a source or feature

1. Add its default path and module flag in `config/default.toml`.
2. Add a source parser and strict schema contract in `ingestion.py`.
3. Add raw/core/mart tables through a new numbered migration. Never edit an
   already-released migration.
4. Materialize the model in `models.py`; carry file lineage and enforce the
   correct grain.
5. Add only curated report output in `reports.py`.
6. Add synthetic tests for schema drift, dedupe and the business rule.
7. Update both architecture and beginner documentation.

Forecast fields must stay forecast-only. Do not join forecast into employee
attendance or correction models.

## Build Windows portable ZIP

The builder needs internet and a build-machine Python with `pip`. The output
does not need Python or internet:

```bash
python packaging/windows/build_portable.py --clean
```

It downloads the official CPython embeddable ZIP, downloads Windows wheels,
extracts them into `runtime/site-packages`, copies the application and produces
`dist/WFMHub-Portable-v0.1.0-win-x64.zip` plus its SHA-256 file.
