# Data organization

This is the **single** data folder used by the tool. It is optional for users who want to keep custom presets and templates.

## Public / shareable (optional)

- `mapping.json`: minimal default mapping template.
- `presets/`: presets and templates (public JSON only). Subfolders are allowed.

## Private / local only

- `proprietary/`: **do not publish**. Local exports and internal filters.
  This folder is excluded by `.gitignore` and should remain private.

## Suggested local layout

```
data/
  presets/
    mapping.template.json
    filters.template.json
    week_view.template.json
  proprietary/          # local-only, ignored by git
    presets/
      my_org.mapping.json
      my_org.filters.json
      my_org.week_view.json
    <your spreadsheets here>
```

## Notes
- Keep local exports in `data/proprietary/` only.
- If you need shareable examples, scrub them and put them under `data/sample/`.
