# how to extract data from website to excel: 5 Working Methods Compared, Plus the Credit Math to Check Before You Pay for a Scraping Tool

Getting website data into Excel comes down to three paths: Excel's own Power Query for simple static tables, a Python or VBA script when you need control, and a scraping API when the site fights back with JavaScript, pagination, or bot protection. This guide walks through each method with its actual steps and breaking points, then shows how a service like ScraperAPI handles the hard cases — including what its plans cost and, more importantly, how its credit system really works before you commit to a subscription.

> **Quick take:** If your target site is Amazon, Google, Walmart, or anything behind Cloudflare, skip the DIY route. ScraperAPI's DataPipeline runs scheduled scrapes from a dashboard and hands you a CSV that opens straight into Excel. Plans start at **$49/month** (100,000 credits), and you can test everything with **5,000 free credits for 7 days, no credit card required**. 👉 [Start the free trial and scrape your first page](https://www.scraperapi.com/?fp_ref=coupons)

## Which method should you use?

Match the method to the site, not to your ambition. The decision looks like this:

| Your situation | Best fit | Why |
| --- | --- | --- |
| Public HTML table, no JavaScript, one page | Excel Power Query | Built into Excel, zero cost, two minutes |
| Multiple pages, custom fields, one-off project | Python script | Free, flexible, but you maintain it |
| JavaScript-heavy site, pagination, or bot protection | Scraping API | Handles rendering, proxies, and CAPTCHAs for you |
| Recurring data collection (prices, rankings) | Scraping API with scheduling | Runs on a schedule without your machine being on |

Most articles ranking for this topic stop at the Power Query tutorial. That's fine as far as it goes, but it leaves out the part that actually determines whether you succeed: what happens when the site doesn't cooperate. The sections below cover both halves.

## Method 1: Excel's built-in "From Web" (Power Query)

For static pages with standard HTML tables, Excel already includes everything you need — no add-ons, no code.

The steps, in Excel's Data ribbon:

1. Go to **Data → Get Data → From Other Sources → From Web**
2. Paste the page URL and click OK
3. In the Navigator window, Excel lists every table it detected on the page — click through the previews until you find your data
4. Press **Load** to drop it into a sheet, or **Transform Data** to clean it up in the Power Query Editor first

This works well when three conditions hold: the data sits in a proper HTML `<table>`, the page serves its content without JavaScript, and you only need one page at a time. Microsoft's own documentation confirms the load-then-shape workflow with Power Query Editor.

The failure modes matter more than the success case. JavaScript-rendered pages come back empty because Power Query fetches raw HTML and never runs the scripts. Paginated listings mean repeating the whole process per page — Power Query can sometimes handle multi-page data if the URL pattern is predictable, but that's a polite way of saying it usually can't. And anything guarded by Cloudflare or similar protection simply won't load. There's also no scheduling; every refresh is manual.

## Method 2: Python or VBA scripts

When Power Query runs out of road, code is the traditional answer. A Python script using `requests` to fetch pages, BeautifulSoup to extract fields, and pandas to write the result with `df.to_excel()` can loop through paginated URLs and store everything as CSV or Excel as it goes. VBA inside Excel can do the same, though maintaining VBA in 2026 is a choice most people only make under duress.

Scripts are free and genuinely flexible. The cost is invisible until later: selectors break when the site redesigns, your IP gets blocked once you query too fast, and JavaScript-heavy sites require a headless browser (Playwright or Selenium) plus proxy rotation to behave like a normal visitor. For a one-off personal project, that maintenance burden is worth accepting. For anything recurring or business-critical, the hours spent patching a broken scraper usually exceed what a scraping service costs.

## Method 3: A scraping API that hands you spreadsheet-ready data

This is where ScraperAPI fits the "website to Excel" problem. It isn't a point-and-click extractor — it's an API plus a no-code dashboard that takes care of proxies, CAPTCHAs, JavaScript rendering, and retries, then returns structured data you can open in a spreadsheet.

The no-code route is **DataPipeline**, and it maps almost one-to-one onto the workflow most people searching this topic actually want:

1. Open DataPipeline in the dashboard and click **New project**, then pick a project type (Amazon product pages, Walmart product, Google Shopping)
2. Under Input settings, choose **List** and paste your ASINs, product IDs, or search queries one per line — or upload a CSV with one item per row for larger lists (up to 100,000 items, no header row)
3. Under Output settings, choose **Download**; for automated setups, enter a Webhook URL so results post to your server as they arrive
4. Pick a scraping frequency: **Just scrape once**, or Daily/Weekly for recurring monitoring
5. Hit **Review & start scraping** — the Job summary panel shows the estimated credit cost before you confirm anything
6. When the run finishes, click **Download results**

The CSV output flattens each result into a row. For Google Shopping, that means `position`, `title`, `price`, `source`, and `extracted_price` become columns — open the file in Excel and it's already a clean table, no parsing needed. Comparing `extracted_price` across runs is how the official price-tracking tutorial suggests spotting price movements by retailer.

Two limits worth knowing before you start. First, recurring schedules require a paid plan — one-time scrapes work on the free tier, but Daily/Weekly monitoring doesn't. Second, results are retained for 30 days, so download them promptly.

Developers can skip the dashboard entirely and call the API directly, with parameters like `output_format` (no extra cost), `render=true` for JavaScript pages, and `autoparse=true` for automatic structured output.

## ScraperAPI plans and pricing (all seven tiers)

Here's the full pricing table as shown on the official pricing page, including the tiers most reviews skip:

| Plan | Monthly price | Annual (per month) | API credits | Concurrency | Get started |
| --- | --- | --- | --- | --- | --- |
| Hobby | $49 | $44.10 | 100,000 | 20 | [ View Hobby plan pricing](https://www.scraperapi.com/pricing?fp_ref=coupons) |
| Startup | $149 | $134.10 | 1,000,000 | 50 | [ View Startup plan pricing](https://www.scraperapi.com/pricing?fp_ref=coupons) |
| Business | $299 | $269.10 | 3,000,000 | 100 | [ View Business plan pricing](https://www.scraperapi.com/pricing?fp_ref=coupons) |
| Scaling (Most popular) | $475 | $427.50 | 5,000,000 | 200 | [ View Scaling plan pricing](https://www.scraperapi.com/pricing?fp_ref=coupons) |
| Professional | $975 | $877.50 | 10,500,000 | 300 | [ View Professional plan pricing](https://www.scraperapi.com/pricing?fp_ref=coupons) |
| Advanced | $1,975 | $1,777.50 | 21,500,000 | 500 | [ View Advanced plan pricing](https://www.scraperapi.com/pricing?fp_ref=coupons) |
| Enterprise | Custom ("Let's talk") | Custom | 22,000,000+ | 500+ | [ View Enterprise options](https://www.scraperapi.com/pricing?fp_ref=coupons) |

Annual billing takes 10% off every plan. Every paid tier includes the same core feature set: JavaScript rendering, premium proxies, JSON auto-parsing, CAPTCHA and anti-bot detection, unlimited bandwidth, and a 99.9% uptime guarantee.

Two plan-level restrictions that don't show up in comparison tables: **Hobby and Startup are limited to US & EU geotargeting only**, and their analytics data is retained for only the last 30 days. Business and above get country-level targeting and unlimited analytics retention.

## What a credit actually costs (the part pricing pages gloss over)

This is where most "ScraperAPI review" articles stop, and where you should keep reading. One credit equals one standard page request — but the site you're scraping and the parameters you send change that number dramatically.

| Request type | Credits per request |
| --- | --- |
| Standard (flat) page | 1 |
| Amazon, Walmart, eBay | 5 |
| Google, Bing (all subdomains) | 25 |
| LinkedIn | 30 |
| Bot-protected sites (Cloudflare, Datadome, PerimeterX, Akamai bypass) | +10 |
| With `render=true` (JavaScript rendering) | +10 |
| With `premium=true` | +10 |
| `premium=true` + `render=true` | 25 |
| `ultra_premium=true` + `render=true` | 75 (paid plans only) |

Run the math on your own use case before picking a plan. The Hobby plan's 100,000 credits is 100,000 ordinary pages — but only 20,000 Amazon product pages, or 4,000 Google searches, or roughly 2,850 Google searches *with* JavaScript rendering (25 + 10 credits each). A plan that looks huge on the pricing card can shrink by a factor of 25 depending on your targets.

Three details soften the blow:

- **Failed requests are free.** ScraperAPI only charges for successful requests (HTTP 200 and 404) and for requests you cancel within 70 seconds. Retries on timeouts cost you nothing.
- **You can check the cost before scraping.** The dashboard's Domain Cost Estimator (also called Domain Multiplier) looks up the credit cost of any URL in advance, the API Playground lets you test target URLs and see actual credit usage, and there's even an API endpoint for it. You can also set a `max_cost` per scrape so a single job can't burn through your balance.
- **`autoparse=true` and `output_format` cost nothing extra** — getting structured, spreadsheet-ready output doesn't inflate your bill.

Every response also includes a `sa-credit-cost` header showing exactly what that request consumed, so there's no guesswork after the fact.

If credits run out mid-month, scraping pauses by default until renewal. Hobby through Professional plans can opt into auto-upgrade, which bumps you to the next plan automatically when credits hit zero; Scaling and above can instead use pay-as-you-go at a fixed per-credit rate, capped at a monthly limit you set.

## Which plan actually fits your project

The plan names suggest a ladder you're supposed to climb. In practice, most use cases land in three buckets:

- **You're testing whether web scraping works for you at all.** Use the free trial first — 5,000 credits, 7 days, no credit card. At standard-page pricing that's 5,000 pages of data; even at Amazon's 5-credit rate, 1,000 product pages is enough to validate a whole workflow. 👉 [Sign up free and check your target site's credit cost](https://www.scraperapi.com/?fp_ref=coupons)
- **You scrape a few hundred to a few thousand pages per month of ordinary sites.** Hobby at $49/month covers it — with one caveat: Hobby is geotargeted to US & EU only, so if you need results from other countries or care about long-term analytics history, jump to Business. If neither applies, don't overbuy; higher tiers mostly buy you concurrency you won't use.
- **You run production monitoring on e-commerce or search data.** Business ($299, 3M credits) is the realistic entry point — country-level geotargeting and unlimited analytics retention start there. Scaling adds pay-as-you-go headroom, which matters when a monthly crawl volume spikes unpredictably.

One more asymmetry worth flagging: concurrency, not credits, is what stalls big jobs. Exceeding your plan's thread limit returns a 429 error and the request is rejected rather than queued. If your pipeline needs 300 parallel requests, no amount of credits on a Hobby plan will help — you'd need Professional territory. For a single DataPipeline job, though, even 20 concurrent threads on Hobby processes a few thousand pages in reasonable time.

## FAQ

**How much does it cost to scrape Amazon or Google specifically?**
Amazon costs 5 credits per request, Google and Bing cost 25, LinkedIn 30, and a standard page costs 1. These multipliers are the single biggest factor in your real monthly bill — a "100,000 requests" plan is really a "20,000 Amazon pages" plan.

**Do I get charged when a scrape fails?**
No. Only successful requests (200 and 404 status codes) and requests you cancel within 70 seconds are billed. Everything else — timeouts, retries, blocked attempts — is free.

**How long do I keep access to my scraped results?**
DataPipeline results are retained for 30 days, so download them into Excel soon after each run finishes.

**Is there a free option to try before paying?**
Yes. A 7-day trial with 5,000 credits requires no credit card, and a permanent free plan gives 1,000 credits with a maximum of 5 concurrent connections; you can contact support for extra testing credits.

## Before you subscribe

Two things to verify with your own numbers. First, look up your actual target URLs in the Domain Cost Estimator — the difference between 1 credit and 25 credits per request decides which plan is honest value for you, and nobody else's use case transfers. Second, decide whether you need scheduled scraping: if you do, the free tier won't cover it and Hobby is the minimum; if you just need a one-time export, the trial may finish the job entirely at zero cost.

If your targets turn out to be ordinary flat pages at 1 credit each, Hobby's $49/month (or $44.10 on annual billing) will likely outlast your project. If you're tracking Amazon or Google Shopping prices daily, start your estimate from Business. 👉 [Check current plans and credit costs on ScraperAPI](https://www.scraperapi.com/pricing?fp_ref=coupons)
