Automated MTG Proxy Print Workflow
This repository automates the process from a text card list to printable, duplex-aligned proxy sheets:
Download card art with MTG-Art-Downloader.
Render cards with Proxyshop and Adobe Photoshop.
Assemble landscape US Letter sheets with fronts and matching backs.
Pause for PDF review.
Open the system print dialog for final printing.
The tools and card artwork are not included. Use only artwork you are permitted to use.
Requirements
Windows 10 or 11
PowerShell 5.1 or newer
Python 3.10–3.12
Poetry
Adobe Photoshop desktop
A source checkout of Proxyshop 1.13.2 for headless rendering
A source checkout of MTG-Art-Downloader
Python PDF/image packages:
```powershell
python -m pip install pillow numpy reportlab pypdf
```
Optional:
SumatraPDF, for a cleaner print-dialog workflow
Epson ET-8550 or another duplex-capable printer
Repository files
Place these files together, or adjust the paths when running the orchestrator:
```text
orchestrate_proxy_workflow.ps1
assemble_card_sheets.py
horizontal_cardBack_3x3.pdf
```
MTG-Art-Downloader and Proxyshop remain separate checkouts.
Install
```powershell
# MTG-Art-Downloader
cd path\to\MTG-Art-Downloader
poetry install --no-root

# Proxyshop source checkout
cd path\to\Proxyshop
poetry install --no-root
```
Confirm that Photoshop is installed, starts normally, and can be automated on the Windows account running the workflow.
Card list
Create `cards.txt` in the MTG-Art-Downloader directory:
```text
4 Beza, the Bounding Spring (BLB) 287
2 Elves of Deep Shadow (RAV) 161
Lightning Bolt
```
A leading number sets the quantity.
A line without a quantity produces one copy.
Set codes and collector numbers may be included for the downloader.
The sheet assembler matches cards by name and ignores unlisted image files.
Run
From this repository:
```powershell
powershell -ExecutionPolicy Bypass -File .\orchestrate_proxy_workflow.ps1 `
  -DownloaderDir "path\to\MTG-Art-Downloader" `
  -ProxyshopDir "path\to\Proxyshop" `
  -AssemblerScript ".\assemble_card_sheets.py" `
  -BackPdf ".\horizontal_cardBack_3x3.pdf" `
  -ProxyshopMode SourceHeadless `
  -DuplexBinding long-edge
```
Use `-ProxyshopMode GuiManual` if you prefer to click Render All in Proxyshop yourself.
During the run:
Missing downloads or failed renders pause the workflow.
Type `c` to continue or `q` to stop.
Review the generated PDFs before approving printing.
In the system print dialog, select the correct paper, quality, and duplex preset.
Output
Generated files are placed under:
```text
Proxyshop\orchestrated_runs\<timestamp>\
```
The assembler produces one merged PDF for every group of up to eight cards. Each PDF contains:
Landscape US Letter pages
Eight 63×88 mm cards per sheet
An embedded 3 mm black bleed around each card
Matching back pages using the selected duplex alignment
Duplex alignment
Use the binding mode that matches the printer's automatic duplex behavior:
```powershell
-DuplexBinding long-edge   # default
-DuplexBinding short-edge
-DuplexBinding none
```
If a test sheet places backs under the wrong fronts, run one sheet again with the other binding mode. Always verify with plain paper before using card stock.
Troubleshooting
Missing Python packages in a Poetry project: run `poetry install --no-root` in that project.
Proxyshop 1.13.2 on Python 3.10: the orchestrator applies small compatibility fixes for `NotRequired` and `ForwardRef`. The original `src\layouts.py` is saved as `src\layouts.py.orchestrator-backup`.
No rendered cards: check Proxyshop's `art`, `out`, and `logs` directories.
Missing listed cards: review MTG-Art-Downloader's `failed.txt`, then continue only if skipping those cards is acceptable.
Incorrect card size or position: confirm the PDF is printed at 100% scale with no fit-to-page option.
Epson preset automation: use the default manual system-dialog mode unless you have tested the experimental preset mode.
Important
Do not enable unattended printing until a full low-cost test run has been reviewed. The orchestrator intentionally stops for review before printing.
