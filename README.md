# Power BI Labs — Five Industries, Six Labs Each

Thirty hands-on Power BI labs across five industries, each series building
from a messy Excel export to a published, secured, production-shaped report.
Same arc every time, different data, different problems, because the
problems are what change between real projects, not the menu clicks.

## Why five industries instead of one deep series

One dataset teaches one dataset's problems. A three-location coffee chain's
data quality issues are not a hospital's, and neither look like a public
library's circulation history. Working through five forces the same
underlying skills, star schemas, DAX measures, RLS, what-if analysis, to
show up in genuinely different shapes, which is closer to what switching
between real clients or projects actually feels like.

## The industries

| Industry | Data | License | Real or synthetic |
|---|---|---|---|
| [Coffee shop](labs/coffee-shop/) | Maven Roasters coffee shop sales, 3 NYC stores | CC0 (Kaggle) | Real, date range extended synthetically for time-intelligence labs |
| [University](labs/university/) | Enrollment and library-usage records | — | Fully synthetic, generated in Lab 01, deliberately carries no PII |
| [Healthcare](labs/healthcare/) | Hospital appointment operations + NHS England A&E benchmark | OGL v3.0 (NHS data) | Synthetic appointments, real NHS trust-level benchmark data |
| [Libraries](labs/libraries/) | Seattle Public Library checkouts by title | Public / open (data.gov) | Real |
| [Retail](labs/retail/) | UCI Online Retail, UK e-commerce transactions | CC BY 4.0 | Real, with a synthetic inventory table and, from Lab 06, a synthetic second sales channel |

## The six labs, every time

| # | Lab | What it adds | What it breaks |
|---|---|---|---|
| 01 | Excel cleanup | Import, assess, fix, first pivot | Formulas sit beside a pivot and don't scale |
| 02 | Data model | Star schema, first DAX measures | No dashboard, no shared calculation |
| 03 | Dashboard | Interactive report, bookmarks, publish | Refresh depends on one file, one machine |
| 04 | Time intelligence | Period-over-period DAX, a deliberate break-it exercise | No forward-looking view |
| 05 | What-if / RLS | Scenario modeling, row-level security | Access control is hand-maintained, not sourced from a real identity system |
| 06 | Capstone | Everything above, plus one new complication and one genuinely advanced technique | — |

Read a series in order. Each lab assumes the one before it. Lab 06 in every
series is harder than anything before it on purpose, and closes with a
reflection on the whole arc rather than a link to a next lab.

## Working through a lab

Every lab follows the same shape: Objectives, Background, Required
Resources, a Topology diagram, then numbered Parts with Steps. Steps that
could trip you up carry a collapsed **Hint**; every Part ends with a
collapsed **Expected result** so you can check your own work without an
instructor in the room. Both are closed by default on GitHub, click to
open.

Struggling on a step is normal and is most of the point. Open the hint
before you open the DAX reference docs in a new tab. If you're still stuck
after the hint, the **Troubleshooting** table at the end of each lab covers
the most common wrong turns, and the **What Went Wrong When I Did This**
section is an honest account of the mistakes made building these labs in
the first place.

## Setup

- Excel (2021 or later, or Microsoft 365) and Power BI Desktop, both
  already installed. Nothing else to install for any of the five series.
- A free Power BI Service account for the publish steps in Lab 03 onward.
  A Microsoft 365 email works, or a free Fabric trial.
- Internet access is needed only to download real source data: the coffee
  shop, healthcare (NHS benchmark), libraries, and retail series each link
  the exact file to fetch in their own Lab 01. The university series
  generates its data locally in Excel and needs no download at all.
- Two labs use Tabular Editor (free, external tool) for calculation
  groups: the healthcare and retail Lab 06 capstones. Everything else uses
  only Excel and Power BI Desktop.

## Repository layout

```
README.md                 this file

labs/
  coffee-shop/             01-06, real Kaggle sales data
  university/               01-06, synthetic enrollment + library data
  healthcare/               01-06, synthetic appointments + real NHS benchmark
  libraries/                01-06, real Seattle Public Library data
  retail/                   01-06, real UCI Online Retail data
```

## What's deliberately out of scope

- No clinical or diagnosis data anywhere in the healthcare series. It's
  hospital operations only: appointments, wait times, department
  throughput. Stated explicitly in that series' own Lab 01 and restated
  in Lab 06.
- No real person's data anywhere. The university series is synthetic by
  design specifically to avoid enrollment-record privacy questions; the
  real datasets used elsewhere (coffee shop sales, library checkouts,
  online retail transactions) are published open data about transactions,
  not about identifiable individuals.
