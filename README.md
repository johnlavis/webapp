# webapp

Experimental webapps project with plugin marketplace support for agent skills.

## Headroom Proxy

[Headroom](https://github.com/chopratejas/headroom) is a context-compression proxy that reduces LLM token usage by 60-95% with no code changes required.

### Start the proxy

```bash
python3 proxy.py start
```

This installs `headroom-ai[proxy]` automatically on first run, then starts the proxy on port 8787.

Custom port:

```bash
python3 proxy.py start --port 9000
```

### Connect your clients

```bash
# Claude Code
ANTHROPIC_BASE_URL=http://localhost:8787 claude

# OpenAI-compatible clients
OPENAI_BASE_URL=http://localhost:8787/v1 your-app
```

### Manual install

```bash
pip install "headroom-ai[proxy]"
headroom proxy --port 8787
```

## Plugin Marketplace

Skills are installed via `plugin.py`:

```bash
python3 plugin.py marketplace add <owner>/<repo>   # install
python3 plugin.py marketplace list                  # show installed
python3 plugin.py marketplace remove <name>         # uninstall
```

Installed skills are stored in `skills/<name>/` and tracked in `skills.json`.
