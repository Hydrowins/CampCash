CAMPCASH
This is a little project I've been working on to track my college expenses, so feel free to explore around.
Full DIsclosure - UI/UX Wirefframes were done by hand on figma, coding was done by Claude AI, under my own supervision, please dont judge, I am a UI guy who doesnt know how to code to save his life.....
{
  "$schema": "./node_modules/oxlint/configuration_schema.json",
  "plugins": ["react", "typescript", "oxc"],
  "options": {
    "typeAware": true
  },
  "rules": {
    "react/rules-of-hooks": "error",
    "react/only-export-components": ["warn", { "allowConstantExport": true }]
  }
}
```

See the [Oxlint rules documentation](https://oxc.rs/docs/guide/usage/linter/rules) for the full list of rules and categories.
