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
| Surveyed by | Sidebar. The phone number is left alone — the template XLOOKUPs it from this name |
| County, State | The alignment sheet's county band, or looked up from the dig's coordinates when no sheet covers it. Correct it in the review table if either is ever wrong |
| Upstream / downstream reference and feet to AGM | The two reference cells, whether they name an AGM or a launch/receive valve |
| Line name | Alignment sheet title block |
| Alignment sheet number | Pipeline number + sheet number off the alignment sheet |
| Tract number, legal description | Alignment sheet PLSS band, matched to the dig's station |
| HCA | Alignment sheet HCA band, falling back to the dig sheet's Is HCA column |
| Directions | Mapbox, using the same logic as the Dig Site Directions Generator |
| Aerial image | Satellite tiles + the uploaded KMZ centerline and placemarks |
| Dig position | The dig sheet's ILI coordinate, or measured along the KMZ centerline between the two referenced AGMs when the sheet carries none |
| Cheat sheet weld rows | Weld distances, with formula signs from the line's stationing direction |
| Cheat sheet flavour | Whichever uploaded template matches the dig sheets' stationing direction |

Elevation, EDOC, survey date and the pre-dig photos are field measurements, and
the staking report's latitude and longitude are too - the app never writes them.
The cheat sheet's Latitude and Longitude columns are filled from the dig sheet
(anomaly row). The Pre_Dig_Photos sheet keeps its own placeholders untouched; only
the StakingReport image slot receives the aerial.

Anything the uploads cannot establish is written as `Unknown` and flagged in
the review table rather than guessed.

## Uploads

1. **Staking report template** (`.xlsm`) — your ONEOK Dig Stake Template
2. **Cheat sheet template(s)** (`.xlsx`) — upload both the ascending and the
   descending version and the app picks the one matching the dig sheets'
   stationing direction. Either one on its own still works; the formulas are
   written per dig regardless
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
cannot round-trip those — it drops DrawingML shapes — so the XLSMPATCH section
edits the workbook at the package level instead, rewriting only the cells that
change and repointing the report's image slot at the aerial. Everything else,
`vbaProject.bin` included, comes through byte-identical.

The aerial is added as a **new** media part and only the StakingReport
picture's own relationship is moved to it. Overwriting the bytes of the part
already in the slot would be simpler, but in this template that part
(`image4.jpeg`) is shared with four placeholders on the **Pre_Dig_Photos**
sheet — so the aerial replaced those too. Pre-dig photos are taken in the field
and pasted in by hand, so those placeholders have to survive untouched.

That also means the aerial image inherits the template's own anchor, so it is
always sized and positioned exactly as the slot in your template. The aerial is
rendered to that slot's actual aspect ratio, read from the template you upload,
so it is never cropped to fit — cropping is what used to clip the title card in
its top-left corner.

Everything drawn on the aerial is sized for the slot's real display size
(roughly 650 px wide), not for the rendered pixel count, so the labels stay
legible after Excel shrinks the image.

That sizing only works with a font that can actually be scaled.
`ImageFont.load_default()` returns a fixed-size bitmap font that ignores the
size it is asked for, so on a host with no system fonts every label came out at
about 11 px however large it was requested — roughly 4 px once Excel shrank the
image. The app now tries the usual DejaVu and Liberation paths, then falls back
to `load_default(size=...)`, which scales (Pillow 10.1+, pinned in
`requirements.txt`). `packages.txt` also apt-installs `fonts-dejavu-core` on
Streamlit Cloud so the first path is normally the one taken.

The workbook is also flagged to recalculate on open. Formula cells carry the
value Excel last computed, and the template's phone number is an XLOOKUP
against the surveyor name with a stale cached result — without the flag the
report opens showing the old value until something nudges the cell.

## Layout

The whole app is one file. That is deliberate: a Streamlit Cloud deploy fails
with a bare `ModuleNotFoundError` if a subpackage does not make it into the
repo, and there is nothing to go missing here.

```
app.py              Everything, in sections:
                      MODELS       Dig record, station and AGM helpers
                      XLSMPATCH    Package-level .xlsm editing
                      DIGSHEET     Dig sheet parsing, xlsx and PDF
                      ALIGNMENT    Title block, PLSS and HCA bands
                      KMZ          Centerline and placemarks
                      AERIAL       Tile fetching and the image render
                      DIRECTIONS   Mapbox, ported from Dig-Site-Directions-Generator
                      CHEATSHEET   Cheat sheet writer
                      STAKING      Staking report writer
                      STREAMLIT UI
requirements.txt
packages.txt        apt packages for Streamlit Cloud (fonts)
.streamlit/config.toml
```

## Notes

- **Tract number** is written as the PLSS section number. That holds on every
  filled report checked (NL3DH-24-F1 → 16, STSB-24-F1 → 6, RJSJ-25-F1 → 7).
  If your tract numbering ever diverges from the section, correct it in the
  review table before generating.
- **PDF cells are read in content-stream order, not by position.** These
  sheets are printed from Excel with text overflowing its columns, so a
  reference cell and the number beside it physically overlap in x. Sorting
  characters by position interleaves them into nonsense
  ("Dan4v9il4le7,." from "Danville," and "4947.55"). Each cell is emitted as
  one contiguous run, so a run ends where x jumps backwards or leaves a gap;
  cells are then assigned to columns by their start position, which the
  overlap does not disturb.
- **References are not always AGMs.** An upstream reference is often
  "LAUNCH VALVE Danville, Sta. 2582+14". The report wants just "Launch Valve",
  so the reference type is taken off the front rather than everything up to
  the first comma.
- **Line name is applied without needing a station match.** It describes the
  pipeline, not a position on it, so any uploaded sheet for that line supplies
  it. Only the sheet number, county, PLSS and HCA bands need the sheet that
  actually covers the dig.
- **Alignment sheet bands are read from the sheet's own structure.** The
  full-width horizontal rules split the sheet into banded rows, each labelled
  down the left edge (COUNTY, PLSS, CLASS, HCA). Within a band, the drawn
  vertical dividers bound each segment, and a segment's station range comes
  from the axis-tick regression — so a station falls in the section it is
  actually drawn in, rather than the nearest label. HCA stretches are thin
  horizontal rules on the HCA row, not filled blocks; when the row is present
  and carries none, that is a reliable "No".
- The station range comes from the axis tick row, not from every
  station-shaped token on the page — real sheets often have no "From X To Y"
  line, and stray tokens produce a nonsense range.
- **Alignment sheets** are matched to a dig by station. When none covers it,
  the app says so and lists every sheet it read with its station range, rather
  than silently leaving the fields Unknown. The sheet number written onto the
  report comes from the filename (`10222_41 Alignment Sheet.pdf` → `10222_41`),
  since the sheet itself prints a zero-padded `Sheet 041`.
- **Stationing direction** is detected from each dig sheet's own odometer and
  stationing columns, so mixed batches across ascending and descending lines
  come out right in one run.
- **Dig names are not always `-F<n>`.** `RBB-26-P1`..`P4` are pigs, and a
  name can carry a letter suffix (`MLT-25-F19B`) that is part of the name —
  the same sheet also holds `MLT-25-F19`, a different dig. The pattern is
  `<prefix>-<yy>-<letters><number><optional letter>`, still guarded against
  running into the feature id printed next to it.
- **Column detection is advisory, not authoritative.** The wrapped header
  merges differently from sheet to sheet, so neighbouring fields land in one
  cell (odometer and stationing together, the weld distances and both AGM
  references swallowed into MOP) or a column is missed entirely. The anchored
  text patterns read the same row correctly because the column *order* is
  fixed even when the header is not readable, so they fill any field whose
  detected cell cannot yield a usable value. A sheet whose columns were read
  cleanly is unaffected.
- **PDF dig sheets**: the anomaly row must satisfy all three conditions at
  once — it starts with a dig name (a value in the leftmost Dig Number column),
  that name matches the heading at the top centre of page one, and the row sits
  in a yellow highlight band. A dig sheet can carry other highlighted rows that
  are not the call anomaly, so no single signal is trusted on its own.
  Field values are then read off the row text with anchored patterns rather
  than by trusting column detection, because the wrapped multi-line headers
  merge unpredictably.
  When nothing matches, the app names which of the three conditions eliminated
  every row and prints what it did see — heading, dig names found, row count,
  highlight bands — so it can be corrected against the real sheet rather than
  guessed at.

## Choosing the cheat sheet template

The ascending and descending cheat sheet templates are identical but for
column C: ascending subtracts the weld distance from the anomaly's station,
descending adds it. The app reads cell C6 to tell them apart, so the choice
does not depend on the file names.

Which one it starts from comes from the dig sheets themselves — each carries
its own stationing direction, and the majority of the batch wins. The station
formulas are still written per dig on top of that, so a mixed batch across
ascending and descending lines comes out right in one run whichever template
was used; picking the matching one just means the sheet also reads the way
that line's crew expects. If only one template is uploaded, that one is used
and the app says so.

## When a sheet carries no coordinates

Some dig sheets print ILI Latitude/Longitude only on the AGM rows, never on the
anomaly row (the RBB and MLT sheets are like this). The sheet still says how far
the dig sits from the AGM either side of it, and the KMZ carries both those AGMs
and the centerline, so the position is **measured along the line** rather than
guessed:

- both referenced AGMs are located on the centerline;
- the dig is placed between them, in proportion to the sheet's own two
  distances — so a KMZ whose centerline length disagrees slightly with the
  odometer still lands the dig in the right relative spot;
- the direction comes from where the two AGMs actually fall on the line, not
  from an assumption about how the KMZ was drawn.

If only one of the two AGMs is in the KMZ there is nothing to set the direction
from, so the dig is **not** placed — a stake on the wrong side of an AGM is
worse than no aerial. The app says which digs it placed, and flags it when the
KMZ and the dig sheet disagree about how far apart the AGMs are, or when the
AGMs sit well off the centerline (usually the wrong line's KMZ).

A position worked out this way feeds the aerial image, the driving directions
and the county lookup. It is never written into the report's Latitude and
Longitude cells — those stay blank for the field measurement, as always.

## County and State

They come off the alignment sheet's county band when a sheet covers the dig.
That always wins; it is the document the report is checked against.

When no alignment sheet covers it, they are looked up from the dig's
coordinates instead — FCC first, then the Census geocoder, then Mapbox, so one
outage does not leave the field blank. Results are cached by rounded position,
so a batch along one line makes one request rather than twenty. If every
service fails the field stays `Unknown`; nothing is guessed. The lookup can be
turned off in the sidebar.

Note that the geocoders return a full state name (`Oklahoma`) while the
alignment sheets print the abbreviation (`KS`).

County and State are written on every report, `Unknown` included. Leaving the
cell untouched would keep whatever the template shipped with — the previous
line's county — and a wrong county in a signed report is worse than a blank
one.
