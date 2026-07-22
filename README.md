# EuGo-docint

Stateless, cluster-internal document-understanding service for the EuGo platform:
files in → Markdown, tables with typed numeric cells, image descriptions, per-file
warnings/errors out. Document-level language only — no compliance semantics, no storage,
no document content in logs.

- Design (authoritative): [docs/superpowers/specs/2026-07-19-eugo-docint-design.md](docs/superpowers/specs/2026-07-19-eugo-docint-design.md)
- Conventions & commands: [CLAUDE.md](CLAUDE.md)

## API

`POST /v1/extract` — `multipart/form-data`: N file parts named `files` (pdf/docx/pptx/html/xlsx/jpg/png),
optional `hints` part `{"<filename>":{"purpose":"bom|photo"}}`. Well-formed requests always return
`200` with per-file success or error. Also: `GET /healthz`, `GET /info`, OpenAPI JSON in Development.

## Run

```bash
dotnet run --project src/AppHost      # Aspire dashboard + telemetry (dev)
dotnet run --project src/DocInt.Api   # plain service on http://localhost:8090
```

Azure engines activate when configured (user-secrets or env; endpoint without key → DefaultAzureCredential):
`DocumentIntelligence:Endpoint`, `AzureOpenAI:Endpoint`, `AzureOpenAI:DeploymentNameVision` (default `gpt-4.1-mini`).
Unconfigured engines answer with per-file `engine_unconfigured` — the service always boots.

## Test

```bash
dotnet restore src/DocInt.slnx
dotnet build --no-restore src/DocInt.slnx
dotnet test --no-build src/DocInt.slnx
```

Live smoke against real Azure is env-gated — see CLAUDE.md. Container: `docker build -t eugo-docint .`
(CI builds linux/amd64 + linux/arm64). Kubernetes deployment lives in the EuGo-infra repo.
