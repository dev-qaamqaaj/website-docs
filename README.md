# OBD AI Solutions: Client Onboarding Kit

A branded onboarding kit (white, burgundy and black) for web-development clients. It includes a kickoff deck, a welcome pack, an onboarding form, the service agreement, a payment schedule and a go-live handover certificate. Everything is generated from **one config file**, so for the next client you edit `config.json` and rebuild.

## What's in `dist/`

| File | Purpose | Who signs |
|---|---|---|
| `00 - Project Kickoff Deck.pptx / .pdf` | 13-slide deck to present on the onboarding call | – |
| `01 - Welcome & Onboarding Pack.pdf` | Founder letter, credentials, project status, roadmap, pricing, support | – |
| `02 - Client Onboarding Form.pdf` | Client details, CNIC/NTN, domain choices, mailboxes, package, declaration | Client |
| `03 - Website Service Agreement.pdf` | The contract (18 clauses + Schedules A: scope, B: fees, C: support SLA) | Both |
| `04 - Fee & Payment Schedule.pdf` | Instalments, third-party costs, bank details, payment record | Both |
| `05 - Go-Live Acceptance & Handover.pdf` | Delivery checklist, credentials handover, acceptance sign-off | Both (at launch) |

## Rebuild

```bash
pip install python-pptx jinja2
python3 build.py            # needs Chromium/Chrome; LibreOffice Impress for the deck PDF
python3 build.py --no-deck  # documents only
```

- **Logo:** drop your logo into `assets/logo/` as `obd-logo.svg` or `obd-logo.png` and rebuild. Until then, a built-in "OBD" monogram is used.
- **Fonts:** Playfair Display + Inter (OFL) live in `assets/fonts/`. Install them on the machine you present from so the PPTX looks the same as the PDF.
- **Blanks:** any empty value in `config.json` (client name, CNIC, city, OBD address/NTN, and so on) prints as a fill-in line.

## Fill in before sending

- `client.name`, `client.business`, `client.email`, `client.address`
- `company.email`, `company.phone`, `company.website`, `company.address`, `company.ntn`
- `project.jurisdiction_city` (arbitration seat / courts)
- Bank / wallet details: add them in `templates/04_payment_schedule.html`, or write them in by hand

> These templates are a solid commercial starting point, but they are not legal advice. Have a lawyer or tax advisor in Pakistan review the agreement once, especially the tax and dispute clauses, before you use it widely.
