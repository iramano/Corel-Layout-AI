# Corel AI

Experimental automation system for intelligent CorelDRAW layout adaptation using VBA and AI-assisted visual analysis.

The project was created to automate a common graphic design production task:

> adapting existing artwork to new dimensions while preserving visual hierarchy, product proportions, prices, logos, typography and campaign identity.

Instead of simply scaling an entire design, Corel AI analyzes the structure of the artwork and generates layout instructions that reposition and resize individual elements.

---

## What it does

Corel AI can:

- inspect the structure of a CorelDRAW document;
- export object geometry and hierarchy;
- export visual previews for AI analysis;
- interpret natural-language resize requests;
- generate layout instructions;
- preserve original artwork while creating adapted pages;
- reposition individual objects;
- resize products, prices, logos and text independently;
- stretch suitable backgrounds and geometric elements;
- maintain visual relationships between elements;
- detect objects outside the page;
- apply semantic layout rules;
- process multiple artworks in batch mode.

---

## Batch processing

Multiple artworks can be placed in the same CorelDRAW document.

Each artwork is grouped and treated as an independent composition.

Example:

```text
Artwork 1
Artwork 2
Artwork 3
Artwork 4
...

The system exports a batch scene and preview:

cena_lote.txt
preview_lote.png
pedido_lote.txt

AI analysis then generates:

layout_lote_01.txt
layout_lote_02.txt
layout_lote_03.txt
...
CorelDRAW automatically creates one new page for each artwork and applies the generated layout.

## Architecture

CorelDRAW
    ↓
VBA scene exporter
    ↓
cena.txt / cena_lote.txt
    +
preview.png / preview_lote.png
    +
natural-language request
    ↓
AI layout analysis
    ↓
layout instructions
    ↓
VBA layout executor
    ↓
new CorelDRAW page

...

## Scene analysis

The exporter can identify:
- object type;
- object hierarchy;
- groups;
- text;
- font information;
- position;
- width;
- height;
- fill;
- outline;
- nested objects.
Hierarchical object IDs are represented as:

1
1.1
1.2
1.2.1
1.2.2

## Layout commands

The layout engine currently supports commands such as:

PAGE
FILLPAGE
OBJECT
SET
STRETCH
FITINSIDE
SCALE
ALIGN
ABOVE
BELOW
RIGHTOF
MARGIN
KEEPTOGETHER
DISTRIBUTE
ZONE
PLACEZONE
GAP
SAFEAREA
AUTOFIT
NOCOLLIDE

Example: 

PAGE|123|13.5
FILLPAGE|1.4
OBJECT|1.1|12|6.75|18.2
OBJECT|1.2.2|46|6.75|40
OBJECT|1.2.1|85.5|6.75|27
OBJECT|1.3|112|10.5|14

## Semantic analysis

The system also supports semantic roles such as:

PRODUCT
TITLE
LOGO
PRICE_PRIMARY
PRICE_SECONDARY
CAMPAIGN
LEGAL
BACKGROUND
DECORATION

Artwork-specific semantic blocks can also be created:

BRAND
PRODUCT
OFFER
FOOTER

This allows the layout engine to reason about groups of objects instead of treating every object as an unrelated shape.

## Production workflow

Typical single-artwork workflow:

Prepare artwork
    ↓
Export scene
    ↓
Export preview
    ↓
Describe requested format
    ↓
AI generates layout
    ↓
Apply layout
    ↓
Review

...

Batch workflow:

Group multiple artworks
    ↓
Select all artwork groups
    ↓
PrepareLote
    ↓
AI generates multiple layouts
    ↓
AplicarLote
    ↓
One adapted page per artwork

...

## Project status

This project is experimental and under active development.
Current focus:
- robust production batch processing;
- visual hierarchy preservation;
- automatic composition evaluation;
- collision prevention;
- semantic grouping;
- layout occupancy optimization.

## Requirements

- CorelDRAW with VBA support
- Microsoft Visual Basic for Applications
- Windows
- AI assistant capable of reading the exported scene and preview files
The current development environment uses CorelDRAW and VBA.

## Repository structure

corel-ai/
├── VBA/
│   └── Modulo11.bas
├── docs/
│   ├── CONTEXTO_PROJETO.md
│   ├── INSTRUCOES_ROLES.txt
│   ├── INSTRUCOES_BLOCKS.txt
│   └── INSTRUCOES_REVISAO.txt
├── README.md
├── LICENSE
└── .gitignore

Generated production files are intentionally excluded from source control.

## Disclaimer

Corel AI is an independent experimental project.
It is not affiliated with, endorsed by, sponsored by, or associated with Corel Corporation or Alludo.
CorelDRAW is a trademark of its respective owner.
No client artwork or proprietary production assets are required to use the project.









