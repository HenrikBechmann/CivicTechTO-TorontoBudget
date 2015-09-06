# Better Taxonomy Project

By Henrik Bechmann, July 2015; contents first created February, 2015

# Summary

I’ve assembled a dataset (with some difficulty) of the City of Toronto Budget Summaries going back to 2003 (I couldn’t quite get back to amalgamation – 1998 – in this round). Here's a [sample](https://drive.google.com/open?id=0B208oCU9D8OuNnlIbVVSdUxoYms). For presentation and analysis, I’ve created a better taxonomy for the detailed line items of the summaries, and using automation (Google App Script) created a fairly large set of time series for the data. Finally I’ve offered a preliminary analysis of the data.

My research report can be downloaded [here](https://drive.google.com/open?id=0B208oCU9D8Ouc0JvVERXVWVsRW8&authuser=0), and the original google spreadsheets can be seen [here](https://drive.google.com/open?id=1R5B_HMmDISyCfZcxS34sYY6iReGFbMCF1olzBJ7N9zU&authuser=0). Also available in this github folder is the source code that creates the time series and graphs [here](https://github.com/HenrikBechmann/CivicTechTO-TorontoBudget/blob/master/bettertaxonomy/sourcecodecopyfeb24.gs).

The report has three things that may be of interest.

First, it offers a more accessible taxonomy for the City of Toronto Operating Budget Summary. So for example rather than ‘Citizen Centred Services “A”‘ or ‘Agencies’ at a high level, it offers what I call program domains of ‘Shared Services’, ‘Citizen Support Services’, and ‘Municipal Services’. The taxonomy rolls the 53 line items of the budget into 10 categories, and then rolls these categories into the aforementioned 3 domains. The introduction of the referenced [report](https://drive.google.com/open?id=0B208oCU9D8Ouc0JvVERXVWVsRW8&authuser=0) provides a Taxonomy at a glance which makes this clear.

Second, it applies this taxonomy to a time series of Toronto Operating Budgets from 2003 – 2015. The time series includes many variations of tables and graphs, including notably an inflation adjusted series (based on the Bank of Canada inflation calculator). The tables and graphs are automatically generated based on the annual budget summary tabs so that changes to the model can be made realistically.

Third, it offers a preliminary analysis of these time series, notably that the Toronto budget has gone up about 35% since 2003 on an inflation adjusted basis.

So there are lots of avenues available for investigation.

As mentioned above, the sheets themselves are available on google docs [here](https://docs.google.com/spreadsheets/d/1R5B_HMmDISyCfZcxS34sYY6iReGFbMCF1olzBJ7N9zU/edit?usp=sharing), and if anyone wants the source code for the automated table and chart generation, it's available here: [sourcecodecopyfeb24.gs](https://github.com/HenrikBechmann/CivicTechTO-TorontoBudget/blob/master/bettertaxonomy/sourcecodecopyfeb24.gs).

The brief background section in the [report](https://drive.google.com/open?id=0B208oCU9D8Ouc0JvVERXVWVsRW8&authuser=0) explains my history and interest in this.

# Spreadsheet tabs

Here's the breakdown of the [spreadsheet](https://docs.google.com/spreadsheets/d/1R5B_HMmDISyCfZcxS34sYY6iReGFbMCF1olzBJ7N9zU/edit#gid=931374146) tabs. Note that these sheets are all created by my automation code. Some of the titles could be improved.

**Consolidated Actual**: (the one you used for the visualization I think): Nominal (actual) dollars, 2003 to 2015

3 sections: base data, then rolled up to categories, then categories rolled up to domains.

This is the basis for all other time series.

**Period Changes**: first and last years compared.

3 sections: actual (nominal dollars); constant (inflation adjusted) dollars; common (common size analysis) percents. See the qualifiers in the 'change' columns for selection.

Each section with categories, then rolled up to domains.

**Actual Reference Charts**: The most recent year only (obviously actual)

categories, rolled up to domains.

**Actual Trend Charts**: The time series, actual dollars

categories, rolled up to domains

[**Actual Trend Charts Transposed**: (internal use only)]

**Constant Trend Charts**: The time series, in constant (inflation adjusted) dollars

categories, rolled up to domains

[**Constant Trend Charts Transposed**: (internal use only)]

**Common Trend Charts**: Common size analysis, by year. 

categories, rolled up to domains

[**Common Trend Charts Transposed**: (internal use only)]

The rest of the tabs are for internal use -- the input data for the automated output (the above tabs).
