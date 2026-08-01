# Power BI dashboard recipe

The Router keeps everything in SharePoint lists, so Power BI connects
natively — no exports, no gateway, no backend. This page is the build
recipe for whoever owns Power BI at your agency; the first dashboard is
about an hour of work.

## What you get

- **Outstanding-response scorecard** — which divisions still owe comments
  per bill, overdue flagged against the due date.
- **Fiscal totals** — estimated cost by fiscal year, stacked by division,
  sliced by impact severity (the running total behind budget requests).
- **Responsiveness** — average hours from post to "Commented," per
  division, against your response window.
- **Session burn-down** — bills published vs. fully commented over time.

## Prerequisites

- Power BI Desktop (free) to build; a Pro license (typically included in
  government M365 G-SKUs) to publish and share. Available in GCC.
- Read access to the Router's SharePoint site and lists.
- The add-in maintains a **`CostValue`** number column on the BillTracker —
  the numeric twin of the human-typed `EstimatedCost` text ("$1.2M",
  "(50,000) savings"). Harvest writes it automatically; **Load fiscal
  rollup** backfills it for hand-typed rows. Run one rollup load before
  your first refresh so the column is populated.

## Connect (10 minutes)

1. Power BI Desktop → **Get Data → SharePoint Online List** → your site
   URL (e.g. `https://tenant.sharepoint.com/sites/Legislative`) →
   **2.0** implementation.
2. Select three lists: your **bill tracker** (default `BillTracker`),
   **audit** (`LegislativeAudit`), and **routing** (`LegislativeRoutingMatrix`).
3. In Power Query, keep these columns and remove the rest:
   - BillTracker: `Title` (bill), `Division`, `Status`, `DueDate`,
     `FiscalYear`, `EstimatedCost`, `CostValue`, `ImpactSeverity`, `Created`, `Modified`
   - LegislativeAudit: `Title` (bill), `Status`, `Divisions`, `Created`, `PublishedBy`, `SourceSubject`
   - LegislativeRoutingMatrix: `Title` (code), `DivisionName`, `Emails`
4. Set types: `DueDate`/`Created`/`Modified` → Date/Time, `CostValue` → Decimal.

### Fallback cost parser (only for rows CostValue hasn't reached)

Add a custom column to BillTracker if you see blank `CostValue` on rows
with text costs:

```
= let t = Text.Lower(Text.Trim([EstimatedCost] ?? "")),
      neg = Text.StartsWith(t, "(") or Text.StartsWith(t, "-"),
      clean = Text.Select(t, {"0".."9", ".", "k", "m", "b"}),
      num = try Number.From(Text.Select(clean, {"0".."9", "."})) otherwise null,
      mult = if Text.Contains(clean, "b") then 1000000000
             else if Text.Contains(clean, "m") then 1000000
             else if Text.Contains(clean, "k") then 1000 else 1
  in if num = null then null else (if neg then -1 else 1) * num * mult
```

## Model (5 minutes)

- Relate `BillTracker[Division]` → `LegislativeRoutingMatrix[Title]`
  (many-to-one) so visuals can show full division names.
- Relate `BillTracker[Title]` → `LegislativeAudit[Title]` (many-to-many is
  fine here; the audit adds publish timestamps and who published).

## Measures (copy-paste DAX)

```
Total Estimated Cost = SUM ( BillTracker[CostValue] )

Outstanding = CALCULATE ( COUNTROWS ( BillTracker ),
    BillTracker[Status] IN { "Pending review", "In review" } )

Overdue = CALCULATE ( COUNTROWS ( BillTracker ),
    BillTracker[Status] IN { "Pending review", "In review" },
    BillTracker[DueDate] < NOW () )

Avg Response Hours = AVERAGEX (
    FILTER ( BillTracker, BillTracker[Status] = "Commented" ),
    DATEDIFF ( BillTracker[Created], BillTracker[Modified], HOUR ) )

Awaiting Estimates = CALCULATE ( COUNTROWS ( BillTracker ),
    ISBLANK ( BillTracker[CostValue] ),
    NOT BillTracker[Status] IN { "No comment needed" } )
```

`Avg Response Hours` uses row Created→Modified as a proxy for
post-to-comment time; it reads slightly high if rows are edited after
commenting, which is acceptable for a trend line.

## Suggested pages

1. **Today** — cards: Outstanding, Overdue (red), Total Estimated Cost
   for the current SFY; table of overdue rows (bill, division, due date).
2. **Fiscal** — stacked column: Total Estimated Cost by `FiscalYear` and
   `Division`; matrix: FiscalYear × ImpactSeverity with Awaiting
   Estimates alongside so "TBD" rows stay visible, never silently zero.
3. **Divisions** — bar: Avg Response Hours by division vs. a constant
   line at your response window (48h); Outstanding by division.
4. **Session** — line: cumulative bills published (audit Created) vs.
   cumulative fully-commented over time.

## Publish

Publish to your workspace → schedule refresh (SharePoint credentials =
your org account; daily is plenty, hourly during session) → in the
legislative Teams channel, **+ Add a tab → Power BI** and pick the
report. Divisions then see the scorecard right next to the BillTracker
tab where they mark their status.

> Privacy note: the dashboard reads the same lists your team already
> has access to, under the viewer's own permissions. Nothing new is
> collected and nothing leaves your tenant.
