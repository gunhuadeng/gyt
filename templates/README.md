# Private Word template directory

The production report template and its field baseline are intentionally ignored by Git because they can contain:

- property owner names
- certificate numbers
- addresses
- signature images
- identity and property-right documents

For local use, place these private files here:

```text
templates/formal-report-template.docx
templates/template_baseline.json
```

Without them, the application uses the generic Word draft generator in `office_export.py`.
