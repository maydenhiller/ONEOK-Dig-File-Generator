[README.md](https://github.com/user-attachments/files/31648291/README.md)
# ONEOK Dig File Generator

A Streamlit app that turns dig sheets into filled staking reports and a filled
dig stake cheat sheet.

Upload the templates, drop in a batch of dig sheets, and get back one
`<Dig Name>_Staking Report.xlsm` per dig plus a single cheat sheet covering the
whole batch — with driving directions written in and a correctly sized aerial
image dropped into the report's image slot.

## What it fills in

| Field | Source |
| --- | --- |
| Dig name, ODO, stationing | The dig sheet row carrying a Dig Number |
| Upstream / downstream reference and feet to AGM | US/DS AGM Ref and Distance columns |
| Line name, county, state | Alignment sheet title block |
| Alignment sheet number | Pipeline number + sheet number off the alignment sheet |
| Tract number, legal description | Alignment sheet PLSS band, matched to the dig's station |
| HCA | Alignment sheet HCA band, falling back to the dig sheet's Is HCA column |
| Directions | Mapbox, using the same logic as the Dig Site Directions Generator |
| Aerial image | Satellite tiles + the uploaded KMZ centerline and placemarks |
| Cheat sheet weld rows | Weld distances, with formula signs from the line's stationing direction |

Latitude, longitude, elevation, EDOC, survey date and the pre-dig photos are
field measurements and are deliberately left for manual entry.

Anything the uploads cannot establish is written as `Unknown` and flagged in
the review table rather than guessed.

## Uploads

1. **Staking report template** (`.xlsm`) — your ONEOK Dig Stake Template
2. **Cheat sheet template** (`.xlsx`) — either the ascending or descending
   version; the app writes the right formulas per dig either way
3. **Dig sheets** (`.pdf` or `.xlsx`, several at once)
4. **Alignment sheets** (`.pdf`, optional)
5. **Pipeline KMZ** (`.kmz` or `.kml`, optional — needed for the centerline and
   AGM markers on the aerial image)

## Running it

```bash
pip install -r requirements.txt
streamlit run app.py
```

### Mapbox token

Directions and Mapbox satellite imagery both need a token. Copy
`.streamlit/secrets.toml.example` to `.streamlit/secrets.toml`:

```toml
[mapbox]
token = "pk.your_token_here"
```

On Streamlit Community Cloud, paste the same block into **Settings → Secrets**.

Without a token the app still runs: imagery falls back to Esri World Imagery
(no key required) and the Directions box is left blank for you to fill in.

## How the template is edited

The staking report template carries macros and macro-linked buttons. openpyxl
cannot round-trip those — it drops DrawingML shapes — so `digfiles/xlsmpatch.py`
edits the workbook at the package level instead, rewriting only the cells that
change and swapping the bytes of the image already anchored in the report's
image slot. Everything else, `vbaProject.bin` and `drawing1.xml` included, comes
through byte-identical.

That also means the aerial image inherits the template's own anchor, so it is
always sized and positioned exactly as the slot in your template.

## Layout

```
app.py                  Streamlit UI
digfiles/
  models.py             Dig record and station/AGM helpers
  digsheet.py           Dig sheet parsing, xlsx and PDF
  alignment.py          Alignment sheet title block, PLSS and HCA bands
  kmz.py                KMZ/KML centerline and placemarks
  aerial.py             Tile fetching and the aerial image render
  directions.py         Mapbox directions, ported from Dig-Site-Directions-Generator
  cheatsheet.py         Cheat sheet writer
  staking.py            Staking report writer
  xlsmpatch.py          Package-level .xlsm editing
```

## Notes

- **Tract number** is written as the PLSS section number. That holds on every
  filled report checked (NL3DH-24-F1 → 16, STSB-24-F1 → 6, RJSJ-25-F1 → 7).
  If your tract numbering ever diverges from the section, correct it in the
  review table before generating.
- **Stationing direction** is detected from each dig sheet's own odometer and
  stationing columns, so mixed batches across ascending and descending lines
  come out right in one run.
- **PDF dig sheets** are parsed positionally from the column headers. Excel dig
  sheets are the more reliable input where you have the choice.
