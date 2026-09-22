# Synthetic World Cup Album 2010–2026

Excel, VBA, and Power BI-ready football analytics practice project built with
deterministic synthetic data. It demonstrates workbook design, dashboard
navigation, dimensional modelling, DAX documentation, and reproducible data
preparation without presenting the generated 2026 values as official results.

## What it contains

- An Excel workbook with dashboard, player exploration, tournament comparison,
  match, event, and Power BI preparation sheets.
- CSV tables arranged as a star-schema-style dataset for Power BI practice.
- VBA modules/forms for workbook navigation and player comparison.
- DAX measure examples and a visual specification under `docs/`.

## Open and explore

1. Open the `.xlsx` workbook in desktop Excel.
2. Review the dashboard and navigation links.
3. Inspect the generated CSV tables and relationships in the Power BI sheet.
4. Copy the documented measures from `docs/DAX_MEASURES.md` into a Power BI
   model if desired.
5. Import the `.bas`/`.frm` modules only in a disposable workbook when macros
   are enabled and trusted.

## Data provenance and limits

- The included values are generated for portfolio demonstration and testing.
- The 2026 portion is projected/synthetic and must not be treated as an
  official competition record.
- The repository does not include real customer data, private documents, or
  proprietary code.
- Football statistics, pricing, deployment, and historical accuracy are not
  production or commercial claims.

## Documentation

- [`docs/DATA_DICTIONARY.md`](docs/DATA_DICTIONARY.md) — tables, fields, and
  relationships.
- [`docs/DAX_MEASURES.md`](docs/DAX_MEASURES.md) — documented measure examples.
- [`docs/DEPLOYMENT_GUIDE.md`](docs/DEPLOYMENT_GUIDE.md) — optional Power BI
  Desktop/Service workflow and its external prerequisites.

## License

MIT. See [`LICENSE`](LICENSE).
