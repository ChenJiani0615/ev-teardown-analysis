# ev-teardown-analysis

Agent Skill for automotive electronics teardown benchmarking. Produces four tables — **Mechanical / Material / Cost / Final-Assembly Process** — plus conclusions and a list of supplementary evidence needed.

Decoupled from `vehicle-data-hub-main`; works standalone.

> 🌐 中文说明：[README.md](README.md)

## Repository Layout

```
ev-teardown-analysis/
├── README.md / README.en.md
├── .github/skills/
│   ├── ev-teardown-benchmark/         # Teardown benchmarking (four tables)
│   │   ├── SKILL.md                   # Main workflow (entry point)
│   │   ├── references/                # Output template, evidence rules, part mapping, surface-feature matrix, mechanism library, benchmark sources, architecture discovery
│   │   └── scripts/                   # part_table_check.py BOM self-check / surface_scan.py surface-feature scan
│   └── should-cost-engineering/       # Should-cost analysis (7-sheet workbook, live formulas)
│       ├── SKILL.md
│       └── references/                # Workbook spec, process cost models
├── docs/                              # Plain-language install & usage guide (Chinese)
├── tests/                             # Script tests (fixtures are anonymized demo data)
└── tools/                             # Generic tooling (workbook builders, leaf-list expansion, report export, CS record parsing, etc.)
```

> Teardown photos and sample part lists are **not in this repository** — the whitelist `.gitignore` only uploads skills and generic tooling. Bring your own part list and photos.

## Installation

**Install to your user environment (works across projects)**

```powershell
# Windows
New-Item -ItemType Directory -Force "$env:USERPROFILE\.copilot\skills" | Out-Null
Copy-Item -Recurse -Force ".github\skills\ev-teardown-benchmark" "$env:USERPROFILE\.copilot\skills\"
```

```bash
# macOS / Linux
mkdir -p ~/.copilot/skills
cp -R .github/skills/ev-teardown-benchmark ~/.copilot/skills/
```

> Skills are only discovered in `.github/skills/`, `.agents/skills/`, `.claude/skills/`, or `~/.copilot/skills/`; user-level prompts directories do not host skills.

## Usage

1. Type `/ev-teardown-benchmark` in the chat, or just describe the task (auto-loads when keywords match).
2. Provide: vehicle model / model year / trim + assembly or location + teardown photos or part data.
3. With a CSV/Excel part list, run the self-check first:

```bash
python .github/skills/ev-teardown-benchmark/scripts/part_table_check.py <part-list.csv>
```

The output reports: total mass across part numbers, per-part mass share, material-group summary, dimension-field misalignment warnings, parts with missing fields, and a gap-check hint such as "no power-module / thermal keywords found in the list".

## Example Requests

```
/ev-teardown-benchmark Volvo EX90 Ultra MY2025 rear inverter, part list in parts.csv
/ev-teardown-benchmark OBC of a certain vehicle, full version (per-part deep dive + process step table)
/ev-teardown-benchmark Check field consistency and mass share of this teardown dataset
```

## Extending

| To add… | Edit… |
|---|---|
| Common part classes for new assemblies (BMS, domain controller, thermal) | Section 3 of `references/part-mapping.md` |
| New field traps | Section 2 of `references/evidence-rules.md` |
| New mechanism formulas (e.g. NVH, fatigue) | `references/mechanism-library.md` |
| New surface-feature types or interpretation rules | `references/surface-feature-matrix.md` |
| Output format changes | `references/output-template.md` + Step 5/6 of `SKILL.md` |
| New part-list column names | `FIELD_ALIASES` in `scripts/part_table_check.py` |

## Dependencies

- Required: Python 3.9+ (standard library only)
- Optional: `openpyxl` (for reading .xlsx)

## Design Trade-offs

- **Concise by default**: four tables + conclusions + 5 evidence items. Per-part deep dives and process step tables require an explicit request, to avoid producing an overlong report every time.
- **No invented cost numbers**: without quotes / volume / yield, the cost table only gives cost drivers and formula templates, so it cannot be mistaken for a real cost baseline.
- **Candidate grading**: when evidence is insufficient, A/B/C candidates are listed side by side with the evidence that would "cut between them", instead of guessing a single answer.
- **Degradation checked beyond steady state**: thermal-path conclusions cover 1-D steady state only; when thick TIM, substrate-less devices, large planar interfaces, phase-change materials, or power cycling appear, pump-out / creep / fatigue lifetime mechanisms must be added, with model limitations declared.
- **Per-face surface-feature reading**: the same texture can serve entirely different functions toward different counterparts (coolant / device / TIM / seal / structure / process); writing only one side systematically misses mechanisms.
