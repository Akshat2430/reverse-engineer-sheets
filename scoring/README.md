# Scoring

`sealed.settings.json` is the Claude Code settings file used to run a scoring session blind: it denies reads of the answer-key folder, past transcripts and the repo's `.git`, makes the toolkit and skill read-only, and turns on the macOS sandbox. Permission rules need absolute paths, so replace `/ABSOLUTE/PATH/TO` with the folder that holds this repo before using it:

```
claude -p '/reverse-engineer "samples/<file>"' --settings scoring/sealed.settings.json
```

Start `.venv/bin/python -m sheetre.bridge` in a normal terminal first: inside the sandbox, Excel can only be driven through the bridge. You don't need any of this to use the skill on your own workbooks.
