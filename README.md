# OBD AI Solutions: Client Onboarding Kit

A branded onboarding kit for web-development clients, styled on the OBD logo (violet → indigo on a light, white base). It's currently set up for **GFX-T** (Lahore). It includes a kickoff deck, a welcome pack, an onboarding form, the service agreement, a payment schedule and a go-live handover certificate. Everything is generated from **one config file**, so for the next client you edit `config.json` and rebuild.

## Ready-to-send folder

`GFX-T Onboarding Kit/` holds every document in two formats: `PDF/` to send or print, and `PPT/` for editable PowerPoint versions. `python3 build.py` regenerates it.

## What's in `dist/`

| File | Purpose | Who signs |
|---|---|---|
| `00 - Kickoff Deck.pptx / .pdf` | 11 slides for the onboarding call | – |
| `01 - Welcome Pack.pdf` | Founder letter, project status, launch timeline, pricing, contacts | – |
| `02 - Service Agreement.pdf` | Short professional contract with a fee summary (Pakistani law, Lahore courts) | Both |
| `03 - Go-Live & Handover.pdf` | One-page launch-day checklist and sign-off | Both |

**Current pricing:** development PKR 15,000 (standard 25,000). Deployment is waived to PKR 0 (standard 25,000) and includes 1 configured mailbox. Each additional professional email is PKR 1,500, paid to OBD, plus the mailbox subscription paid by the client to the provider.

## Rebuild

```bash
pip install python-pptx jinja2 pillow   # + LibreOffice Impress for the PPT exports
python3 build.py            # needs Chromium/Chrome; LibreOffice Impress for the deck PDF
python3 build.py --no-deck  # documents only
```

- **Logo:** `assets/logo/obd-mark.png` / `obd-mark-white.png` are transparent cut-outs of the B mark, and the `source-*` files are the originals.
- **Fonts:** Playfair Display + Inter (OFL) live in `assets/fonts/`. Install them on the machine you present from so the PPTX looks the same as the PDF.
- **Blanks:** any empty value in `config.json` (client name, CNIC, city, OBD address/NTN, and so on) prints as a fill-in line.

## Fill in before sending

- `client.name`: who signs for GFX-T
- `company.email`, `company.phone`: OBD contact details (shown on the last slide)
- CNIC/NTN can be written in by hand

> These templates are a solid commercial starting point, but they are not legal advice. Have a lawyer or tax advisor in Pakistan review the agreement once, especially the tax and dispute clauses, before you use it widely.
