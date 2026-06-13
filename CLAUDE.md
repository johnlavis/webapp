# webapp

Experimental webapps project with plugin marketplace support for agent skills.

## Plugin Marketplace

Skills are installed via `plugin.py`:

```bash
python3 plugin.py marketplace add <owner>/<repo>   # install
python3 plugin.py marketplace list                  # show installed
python3 plugin.py marketplace remove <name>         # uninstall
```

Installed skills are stored in `skills/<name>/` and tracked in `skills.json`.

## Installed Skills

@skills/last30days/SKILL.md
