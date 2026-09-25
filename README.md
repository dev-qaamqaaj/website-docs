# OBD AI Solutions: Client Onboarding Kit

A branded onboarding kit for web-development clients, styled on the OBD logo (violet → indigo on a light, white base). It's currently set up for **GFX-T** (Lahore). It includes a kickoff deck, a welcome pack, an onboarding form, the service agreement, a payment schedule and a go-live handover certificate. Everything is generated from **one config file**, so for the next client you edit `config.json` and rebuild.

## What's in `dist/`

| File | Purpose | Who signs |
|---|---|---|
| `00 - Kickoff Deck.pptx / .pdf` | 13 slides to present on the onboarding call | – |
| `01 - Welcome Pack.pdf` | Founder letter, who we are, their 7-page site, go-live plan, pricing, contacts | – |
| `02 - Onboarding Form.pdf` | Short form, pre-filled with GFX-T's details | Client |
| `03 - Service Agreement.pdf` | Plain-language contract (Pakistani law, Lahore courts) with a fees & package page | Both |
| `04 - Go-Live & Handover.pdf` | One-page launch-day checklist and sign-off | Both |

## Rebuild

```bash
pip install python-pptx jinja2 pillow
python3 build.py            # needs Chromium/Chrome; LibreOffice Impress for the deck PDF
python3 build.py --no-deck  # documents only
```

- **Logo:** `assets/logo/obd-mark.png` / `obd-mark-white.png` are transparent cut-outs of the B mark, and the `source-*` files are the originals.
- **Fonts:** Playfair Display + Inter (OFL) live in `assets/fonts/`. Install them on the machine you present from so the PPTX looks the same as the PDF.
- **Blanks:** any empty value in `config.json` (client name, CNIC, city, OBD address/NTN, and so on) prints as a fill-in line.

## Fill in before sending

- `client.name` / `client.designation`: who signs for GFX-T (CEO or COO)
- `company.email`, `company.phone`: OBD contact details (shown on the last slide)
- CNIC/NTN and bank/wallet details can be written in by hand

> These templates are a solid commercial starting point, but they are not legal advice. Have a lawyer or tax advisor in Pakistan review the agreement once, especially the tax and dispute clauses, before you use it widely.
