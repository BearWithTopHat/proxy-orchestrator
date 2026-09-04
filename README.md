# Automated MTG Proxy Print Workflow

This repository automates the process from a text card list to printable, duplex-aligned 8-card proxy sheets in landscape (This is for the proxyjak paper which recommends that configuration):

1. Download card art with MTG-Art-Downloader.
2. Render cards with Proxyshop and Adobe Photoshop.
3. Assemble landscape US Letter sheets with fronts and matching backs.
4. Pause for PDF review.
5. Open the system print dialog for final printing.

The tools and card artwork are not included. Use only artwork you are permitted to use.

## Requirements

- Windows 10 or 11
- PowerShell 5.1 or newer
- Python 3.10–3.12
- [Poetry](https://python-poetry.org/)
- Adobe Photoshop desktop
- A source checkout of [Proxyshop](https://github.com/Investigamer/Proxyshop/tree/9be3f1a4058f09fd5eedc890ee27f0e665c6b25c) for headless rendering
- A source checkout of [BearWithTopHat/mtg-art-downloader](https://github.com/BearWithTopHat/mtg-art-downloader); use this fork because the original currently has a bug that prevents downloading
- Python PDF/image packages:

```powershell
python -m pip install pillow numpy reportlab pypdf
```

Optional:

- SumatraPDF, for a cleaner print-dialog workflow
- Any duplex-capable printer supported by Windows

## Repository files

Place these files together. The orchestrator also looks for the two project checkouts as subfolders by default:

```text
orchestrate_proxy_workflow.ps1
assemble_card_sheets.py
horizontal_cardBack_3x3.pdf  # not needed for single-sided output
MTG-Art-Downloader\
Proxyshop\
```

Use `-DownloaderDir` and `-ProxyshopDir` when those checkouts live elsewhere.

## Create `horizontal_cardBack_3x3.pdf`

This source PDF is required only for duplex output. Despite its name, it only needs to contain one clean, upright card-back image. The assembler extracts that image and creates the aligned eight-card back pages itself. Use `-SingleSided` with the orchestrator, or `--single-sided` with the assembler, to skip card backs entirely.
You can find a card back image here: https://mtg.fandom.com/wiki/Card_back as an example of the type and size of card back you should use.
### Option 1: Use a print-layout tool

1. Start with a card-back image cropped to the standard 63×88 mm ratio.
2. Place it on a landscape US Letter PDF.
3. Use a card size of 63×88 mm.
4. Export the PDF as `horizontal_cardBack_3x3.pdf`.

### Option 2: Generate it with Python

Save an upright card-back image as `card-back.jpg`. Save the following as `create_card_back_pdf.py`:

```python
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

page_width = 279.4 * mm
page_height = 215.9 * mm
card_width = 63 * mm
card_height = 88 * mm

pdf = canvas.Canvas(
    "horizontal_cardBack_3x3.pdf",
    pagesize=(page_width, page_height),
)
pdf.drawImage(
    ImageReader("card-back.jpg"),
    (page_width - card_width) / 2,
    (page_height - card_height) / 2,
    width=card_width,
    height=card_height,
)
pdf.save()
```

Then run:

```powershell
python create_card_back_pdf.py
```

Place the resulting PDF beside `assemble_card_sheets.py`, or pass its location with `-BackPdf`.

## Install

```powershell
# MTG-Art-Downloader
cd path\to\MTG-Art-Downloader
poetry install --no-root

# Proxyshop source checkout
cd path\to\Proxyshop
poetry install --no-root
```

Confirm that Photoshop is installed, starts normally, and can be automated on the Windows account running the workflow.

## Card list

Create `cards.txt` in the MTG-Art-Downloader directory:

```text
4 Beza, the Bounding Spring (BLB) 287
2 Elves of Deep Shadow (RAV) 161
Lightning Bolt
```

- A leading number sets the quantity.
- A line without a quantity produces one copy.
- Set codes and collector numbers may be included for the downloader.
- The sheet assembler matches cards by name and ignores unlisted image files.

## Run

If you used the default folder layout above:

```powershell
powershell -ExecutionPolicy Bypass -File .\orchestrate_proxy_workflow.ps1
```

For projects in other locations, provide explicit paths:

```powershell
powershell -ExecutionPolicy Bypass -File .\orchestrate_proxy_workflow.ps1 `
  -DownloaderDir "path\to\MTG-Art-Downloader" `
  -ProxyshopDir "path\to\Proxyshop" `
  -AssemblerScript ".\assemble_card_sheets.py" `
  -BackPdf ".\horizontal_cardBack_3x3.pdf" `
  -ProxyshopMode SourceHeadless `
  -DuplexBinding long-edge
```

For front-only sheets with no card backs, add:

```powershell
-SingleSided
```

Single-sided mode does not require `horizontal_cardBack_3x3.pdf` and ignores `-DuplexBinding`.

Use `-ProxyshopMode GuiManual` if you prefer to click **Render All** in Proxyshop yourself. To select a non-default printer or a driver preset, add:

```powershell
-PrinterName "Your Printer Name" -PrintPreset "Your Preset Name"
```

Leave `-PrintPreset` omitted to use the printer defaults.

During the run:

1. Missing downloads or failed renders pause the workflow.
2. Type `c` to continue or `q` to stop.
3. Review the generated PDFs before approving printing.
4. In the system print dialog, select the correct paper, quality, scaling, and siding. Use one-sided/simplex for `-SingleSided` output.

## Output

Generated files are placed under:

```text
Proxyshop\orchestrated_runs\<timestamp>\
```

The assembler produces one PDF for every group of up to eight cards. Each duplex PDF contains:

- Landscape US Letter pages
- Up to eight 63×88 mm cards per front sheet
- An embedded 3 mm black bleed around each card
- Matching back pages using the selected duplex alignment

With `-SingleSided`, each PDF contains only the front sheet for its batch; no back pages are generated or appended.

## Duplex alignment

This section applies only when card backs are enabled. Use the binding mode that matches the printer's automatic duplex behavior:

```powershell
-DuplexBinding long-edge   # default
-DuplexBinding short-edge
-DuplexBinding none
```

If a test sheet places backs under the wrong fronts, run one sheet again with the other binding mode. Always verify with plain paper before using card stock.

## Troubleshooting

- **Missing Python packages in a Poetry project:** run `poetry install --no-root` in that project.
- **Proxyshop 1.13.2 on Python 3.10:** the orchestrator applies small compatibility fixes for `NotRequired` and `ForwardRef`. The original `src\layouts.py` is saved as `src\layouts.py.orchestrator-backup`.
- **No rendered cards:** check Proxyshop's `art`, `out`, and `logs` directories.
- **Missing listed cards:** review MTG-Art-Downloader's `failed.txt`, then continue only if skipping those cards is acceptable.
- **Incorrect card size or position:** confirm the PDF is printed at 100% scale with no fit-to-page option.
- **Printer preset automation:** use the default manual system-dialog mode unless you have tested the experimental preset mode with your printer driver.

## Important

Do not enable unattended printing until a full low-cost test run has been reviewed. The orchestrator intentionally stops for review before printing.
