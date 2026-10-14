# Cellar sheet and website methodology

Updated: 13 October 2026

This guide documents the bottle-level Google Sheets cellar and companion website workflow developed in prior discussions. It preserves reusable methodology without publishing personal inventory, purchases, drinking history, private Sheet URLs or deployment credentials.

## Source of truth and record identity

The native Sheet is the editable source of truth. The website is a published snapshot; changing the Sheet does not automatically update it.

One physical bottle means one row and one website card. Repeated wines remain repeated rows; do not deduplicate them or replace them with a Quantity field. Product identity includes Brand, Model, vintage/edition, release tier, size, packaging and special-edition color. That product key does not uniquely identify a physical bottle.

Inventory holds On Hand and Pre-Arrival bottles. Consumed holds bottles already drunk. Arrival changes stock state; consumption moves one matching physical bottle to Consumed. A separately entered historical consumption does not justify removing an unrelated current bottle. Immutable bottle IDs would be a useful future extension; they are not part of the current schema.

## Sheet schemas

| Inventory column | Meaning and format |
| --- | --- |
| A — Brand | Canonical producer name, preserving accents. |
| B — Model | Cuvée/release tier; Vintage and Plénitude 2 remain distinct. |
| C — Vintage | Text-compatible year or edition, such as `173ème`. |
| D — Packaging | Format first when explicit: `375 mL, Bottle only`. |
| E — Drinkability | `YYYY–YYYY–YYYY`: Early, Prime, Late anchors, en dashes without spaces. |
| F — Status | Calculated Hold, Early, Prime, Late or Past Target. |
| G — Stock | `On Hand` or `Pre-Arrival`. |

The established d’Yquem default is 375 mL, bottle only, unless explicitly overridden. Verify format for other wines from the order or product listing. A promotional photograph containing a box is not proof that packaging was purchased.

Consumed uses Brand, Model, Vintage and Packaging plus the subsequently introduced consumption date and location. No Quantity or freeform Notes column is needed. Use real date cells, ISO `YYYY-MM-DD` in website data, and consistent `DD Mon YYYY` display. Normalize locations consistently, such as venue followed by city. Never fabricate an unknown date or location.

Consumed sorts by consumption date descending, with stable ties. Inventory uses the different drinking-window ordering below.

## Native formatting and safe edits

Read metadata and target cells before writing: values, formulas, notes, validation and formatting. Compare the populated edge with a complete adjacent row. Extend the native Sheet in place, copying safe structure and formatting rather than rebuilding through an export.

Preserve cell dimensions, font family/size, alignment, number formats, borders, frozen headers and conditional formatting. Consumed should match Inventory’s visual language. Preserve native validation or chips when present. Extend applicable table/filter and conditional-format ranges when adding rows. Do not copy another wine’s row-specific values or notes inadvertently.

F2 anchors an array formula. Do not overwrite its spill or move the anchor during sorting. Sort A:E and G together, including their notes, leaving F2 in place. The current Sheet timeline is conditional formatting driven by each row’s E/G; future row-specific formulas would also need preservation. Existing blank formatted rows can be reused; a new row does not require a new worksheet.

## Drinking-window evidence

Windows are cellar-planning estimates, not bottle-condition guarantees or expiry dates. Research the exact cuvée, vintage and release. P1 and P2 are not interchangeable, and packaging alone does not change maturity.

Prefer producer sources for identity and release facts. For maturity, use dated critic evidence with an explicit range, including reputable merchant reproductions. Record critic, publication, review date, cited range, source URL and format limitations in the Drinkability cell note.

Distinguish three quantities:

- **Early:** an evidenced or clearly estimated start of enjoyable drinking.
- **Prime:** one operational target year, usually inferred from a maturity plateau or estimated midpoint; not necessarily a critic’s exact peak.
- **Late:** a stated planning endpoint, sometimes conservative relative to a published maximum; not a spoilage date.

Legacy Champagne records use individually reasoned planning targets and conservative endpoints. Some later additions use the midpoint of a single cited range. These are different estimation methods and must be identified in notes; numerical differences must not be passed off as proof of different ageing potential.

When a midpoint is used, calculate `round((start + end) / 2)`. For example 2020–2032 gives `2020–2026–2032`. Never imply the critic published the midpoint. A review reproduced on a 375 mL listing is not proof that the critic assessed that bottle size. Do not apply an invented universal half-bottle ageing adjustment.

Do not combine the earliest start from one reviewer with the latest end from another without explaining the synthesis. Leave unsupported windows unrecorded rather than inserting zeroes. Revisit stale assessments when a newer comparable tasting materially changes the evidence.

## Maturity classification

For evaluation year `y` and valid ordered anchors `s <= m <= e`:

| Test, in this order | State |
| --- | --- |
| Missing or invalid range | Not recorded |
| `y < s` | Hold |
| `y > e` | Past Target |
| `y == e` | Late |
| `y < m` | Early |
| Otherwise | Prime |

Early includes the start year. Prime includes the middle year and continues until the year before the final year. **Late means the endpoint year only** in this implementation. If anchors coincide, the endpoint test takes precedence.

The current Sheet formula in F2 is:

```excel
=ARRAYFORMULA(IF(E2:E="","",IF(YEAR(TODAY())<VALUE(LEFT(E2:E,4)),"Hold",IF(YEAR(TODAY())>VALUE(RIGHT(E2:E,4)),"Past Target",IF(YEAR(TODAY())=VALUE(RIGHT(E2:E,4)),"Late",IF(YEAR(TODAY())<VALUE(MID(E2:E,6,4)),"Early","Prime"))))))
```

This assumes a valid fixed-width triple. Validate input before writing; the formula does not validate arbitrary text. `TODAY()` follows the Sheet timezone. Collection displays the saved Sheet maturity snapshot; the timeline recomputes maturity for the selected year. The current initial year and Now/reset behavior require deliberate year-rollover maintenance.

## Sorting: Prime, then Late, then Early

The current Inventory and website rule is **Prime ascending → Late ascending → Early ascending**. This supersedes the earlier Early-first ordering. Exact ties remain stable; unavailable windows sort last. Vintage, stock state and maturity label are not the sorting keys.

```javascript
const key = ([early, prime, late]) => [prime, late, early];
```

Re-sort the native values/notes and the static data, and apply the same comparator at website load. Stable ties preserve repeated bottles and the current occurrence-based artist-color mapping. Explicit variant IDs are preferable in a future schema migration.

## Cross-vintage consistency audit

Within each exact Brand + Model group, deduplicate product-level comparisons for analysis only; do not delete physical bottle rows. Compare vintages chronologically against the sorting tuple and separately compare all three anchors. Editions without calendar vintages are a different comparison; singleton groups cannot establish a trend.

An older vintage can correctly receive a later target than a younger one. Accept a reversal only with an evidence-based explanation: structure/acidity, maturity observed at tasting, release/disgorgement history, verified format, or provenance. Art packaging is not an explanation.

Separate **wine-driven exceptions** from **method-driven differences**. Mixed critics, differently dated reviews, conservative cutoffs and midpoint-vs-plateau estimates can create apparent reversals. If evidence supports only broad overlapping windows, do not claim that a two-year inferred Prime difference is scientifically established. Mark it provisional. If the rationale is missing, research and flag it rather than forcing monotonic vintage order.

Record which values are retained, revised or unresolved, with sources and the specific reason. Keep private bottle-level results outside this public repository; publish generalized findings and reusable audit criteria.

## Timeline and color gradients

Collection, Drinking windows and Consumed tabs have consistent text styles and no bottle counts embedded in tab labels. Counts belong in dedicated summaries.

Use a shared timeline axis. The website’s current origin is 2026; its end is at least 2050 and extends to the greatest recorded final year. Use readable 5-year ticks for a short view and 10-year ticks for longer spans. The slider uses the same horizon. A long window must not be falsely depicted as Late in 2050.

| Role | On Hand | Pre-Arrival |
| --- | --- | --- |
| Early | mint `#c7e3d6` | pale blue `#d9e9f8` |
| Prime anchor | gold `#e5b65b` | blue `#94bce0` |
| After Prime | cream `#fce6b9` | light blue `#b9d4ed` |
| Late | coral `#f1b9aa` | coral `#f1b9aa` |

Interpolate around the anchors while keeping years and text labels visible. Color does not replace the state logic or constitute a scientific maturity scale. The Sheet may show a fixed planning viewport while the website exposes a longer axis; both must retain the same full year anchors.

## Bottle imagery and provenance

Use a single, complete bottle: no group photo, box, label-only image, cropped top/base, floor reflection or lifestyle scene. Match producer, cuvée, vintage/edition, release tier, size and artist color. Never silently substitute a different vintage or recolor a different edition.

Search official exact-vintage pages, press assets and technical PDFs first. Use reliable merchants when official material is unavailable. A standardized supplier is useful only when it actually covers the exact variants. Verify the visible year, not just the filename or product title. Group images must not stand in for separate artist-color bottles; blue and pink editions are different assets.

Store provenance in an image manifest: identity key, asset file, source page/original URL, vintage, packaging, dimensions, visible bounds and edit method. The production manifest exists both as JSON and embedded in JavaScript; update both. Current keys join Brand, Model, Vintage and Packaging with `|`, with variant extensions where needed.

Hide source captions and “view wine source” controls on product cards, while retaining the underlying audit metadata. Source credit is not a license; keep third-party photography out of this public methods repository unless separately authorized for redistribution.

## Transparency and half-bottle alignment

Use PNG/WebP with actual alpha transparency. Check on dark and light backgrounds for white fringes, glow, rectangles and reflections. Invisible RGB values in alpha-zero pixels are not a visible backdrop.

Background removal edits a verified real photograph; it must not invent a new label. If AI-assisted, compare year text, foil, geometry and color with the source; reject identity changes. Preserve originals and record editing. Upscaling is not new source detail.

Normalize to visible bottle bounds, not canvas dimensions. The current manifest measures alpha above 20/255. With image size `w,h`, bounds `x,y,x2,y2` and `bh=y2-y`:

```javascript
const bottleHeight = is375mL ? 76.5 : 90;
const imageHeightPercent = bottleHeight * h / bh;
const imageTopPercent = 95 - bottleHeight - bottleHeight * y / bh;
const translateXPercent = -100 * (x + x2) / 2 / w;
```

Preserve aspect ratio; set `left:50%` and apply the horizontal translation. All visible bottle bases land at 95% of their image-area height. Half bottles are 85% of standard displayed height, an editorial visual scale rather than an exact physical measurement. Do not vertically center the smaller bottle: it should share the same bottom baseline. Use the same renderer in all views.

## Website structure and design

The buildless site consists of `index.html`, `style.css`, `app.js`, `data.json`, `consumed.json`, `image-manifest.json` and locally served images. Inventory JSON has a header plus seven-field rows. Consumed objects carry ISO dates and locations. Consumed data and the manifest are also embedded in JavaScript, so changes must keep both representations synchronized.

Use restrained dark styling, consistent typography for words and numbers, clear spacing and a readable wordmark. Avoid a huge cropped bottle as a hero background. Keep complete bottle images aligned. Summaries, filters and static fallback counts must reflect the same data. State clearly that the published site is a snapshot.

## Update and publication workflow

1. Extract exact products, quantities, format and requested stock state from the instruction/order. Read the live Sheet and source before editing; recognize prior successful writes to prevent duplicate retries.
2. Research windows and exact bottle assets, retaining evidence and uncertainty.
3. Add one native row per bottle, preserve formatting/F2, extend applicable rules and sort by Prime → Late → Early. Move notes with records.
4. Read back values, formulas and formatting. Check old-record preservation, added count, stock totals and computed maturity.
5. Derive the website snapshot from verified Sheet values. Update manifests/assets, counts, timeline horizon and cache-version query strings. Do not change unrelated Consumed records.
6. Validate syntax, every image key/file, counts, order, dates and boundary-year logic. Check desktop/mobile appearance, full bottles, alpha edges, half-bottle baseline and the selected-year timeline. Inspect the Google-rendered Sheet; if unavailable, report that limit and inspect native metadata.
7. Package and publish the exact validated source state, preserving the site's audience. Confirm successful deployment, not merely a queued request.
8. Report completed changes and real limitations. Publish only reusable methodology in this public research repository.

The hosted cellar and public research repository are separate. Reuse the site's identity from `.openai/hosting.json`. Through the supported Sites workflow, read access, obtain short-lived credentials in memory, open source, edit/check, then commit/push/package with the source helper. Save the returned commit/archive and deploy the saved version. Use owner-private deployment only when verified owner-only; preserve a shared audience otherwise. Never persist credentials in source or docs.

Sheet and site writes are not transactional. If one succeeds and the other fails, retain the successful work, report partial state and resume only the failed step after fresh reads. Keep the previous deployment available for rollback.

## Final checks

Verify one record per bottle; exact identity/variant; stock separate from maturity; consistent native formats; intact formula spill; cited range and identified Prime estimate; Prime/Late/Early sort; reasoned cross-vintage exceptions; descending consumed dates; no tab counts; correct transparent full-bottle images; smaller half bottles on a shared baseline; full timeline horizon; synchronized data/embedded copies; successful publication; and no private cellar data in public documentation.

## Worked audit examples

These comparisons illustrate how to judge ordering; they are not a list of anyone’s holdings.

**Dom Pérignon P1 2008 versus 2009:** later maturity is well supported. [Decanter’s report of the release sequence](https://www.decanter.com/wine-news/dom-perignon-2009-release-before-2008-369822/) explains that the warmer, more mature 2009 was released before the cooler 2008, which needed more time. That does not by itself prove a universally later final drinking year. Compare like-for-like P1 white wines, not 2009 rosé or P2. Distinguish “needs longer to develop” from “lasts longer.”

**DP 2012 versus 2013/2015:** a reversal can be reasonable. Galloni describes the [2012](https://finewinelibrary.nl/en/product/dom-perignon-vintage-2012) as reticent and concentrated, with a longer range, and the [2013](https://finewinelibrary.nl/en/product/dom-perignon-2013) as open and already enjoyable. His [2015 assessment](https://jeroboamstrade.co.uk/product/2015-dom-perignon-vintage/) also anticipates relatively early maturity. Those descriptions support an earlier-drinking priority for the younger wines; they do not supply exact peak years. Also check critic attribution: the 2012 range 2024–2044 on Fine Wine Library belongs to Audrey Frick/Jeb Dunnuck, not Vinous.

**Cristal 2012 versus 2014:** early approachability is not evidence of shorter lifespan. A [producer interview](https://www.the-buyer.net/tasting/wine/cristal-2014-lecaillon) supports 2014’s early expression. Yet Galloni’s reproduced windows end in [2052 for 2012](https://www.justerinis.com/fine-wines/france/champagne/louis-roederer/cristal-brut/2012) and [2054 for 2014](https://www.vinfolio.com/product/612ec28001000000000550f8/2014-cristal-champagne-louis-roederer). An earlier 2014 planning deadline drawn from a more conservative review must therefore be labelled a policy choice, not consensus that 2012 outlives 2014.

**d’Yquem and midpoint artefacts:** a long review range can place an older vintage after much younger ones. The [2009 merchant reproduction](https://whwc.com/dyquem-sauternes-2009-375ml/) gives Castaing’s 2023–2080 assessment, while his [2021 assessment](https://www.lagunacellar.com/chateau-dyquem-2021) gives 2030–2050. A midpoint calculation preserves that order mechanically; it does not mean the older wine should remain unopened until the midpoint. The [2016 review collection](https://www.farrvintners.com/wine.php?wine=46060) also shows substantial critic disagreement. Keep uncertainty visible rather than describing midpoint arithmetic as biological evidence.
