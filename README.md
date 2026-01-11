# Global Data Analysis in MS Excel

## Introduction

The following project provides an analysis of popular metrics that define coutries around the world. The result of the analysis will be condensed into two interactive dashboards in MS Excel. This project is designed to not only create knowledge from data, but also to give an introduction to advanced data analytics features in Excel.

The following documentation will not cover all steps for you to scrape the data and build those dashboards yourself if you are not experienced in Excel. It is only a short summary to provide context for the project. I strongly recommend to watch this video on YouTube, as it provides the basic knowledge needed for the analysis in Excel: https://www.youtube.com/watch?v=pCJ15nGFgVg

The data used in this project can be found on the following websites:

* https://www.wikipedia.org/
* https://worldpopulationreview.com/
* https://www.numbeo.com/cost-of-living/
* https://ourworldindata.org/grapher/economic-inequality-gini-index
* https://www.economist.com/interactive/big-mac-index
* https://www.worldhappiness.report/

This project is for **educational purposes only**. I do not claim ownership of the underlying datasets. All intellectual property rights belong to their respective creators and providers. Moreover, I cannot guarantee the absolute accuracy or timeliness of the external data used.

## Objective

As mentioned above, the goal of the analysis is, to scrape and clean interesting data from the web and to gain knowledge through interactive dashboards, which make the data accessible to everyone.

Why did I use Excel for this and not a BI-tool like Tableau or Power BI? This project is supposed to teach Excel, because is one of the most used and most useful desktop-applications of all time. Although there are better Applications for this specific purpose, cleaning data and building dashboards in Excel teaches a lot about all sorts of Excel functions that can be used for many use cases.

Before I began collecting data, I drew basic templates for the dashboards in draw.io, to define what I wanted them to look like. Why did I not build the dashboards from scratch in Excel? Two reasons:

1. Being given a task within a specific scope is a common occurrence in most jobs, which makes this project a valuable learning experience for the professional world.
2. Without a goal, one tends to do whatever works easily in the given circumstances without learning much. However, when doing something with a goal, one will make it possible, which may involve changing the circumstances.

For the analysis the following data are of interest:

* **Continent** of each country as a nominal variable that allows filtering
* **Population** as the central metric variable to classify countries
* **GDP** (gross domestic product) and **GNI** (gross national income) per capita to roughly measure income (or rather the average economic output of each person)
* **Gini coefficient** to measure income inequality
* **HDI** (human development index)
* The Economists **Big Mac index** to indicate wheter a country holds an over- or undervalued currency
* Mean years of **life expectancy** and **schooling**
* **Safety**, **health care** and **happieness** to measure quality of life
* ...

The **first dashboard** is supposed to give the user an idea of how countries compare, depending on what is important to the user. Therefore, there are sliders on the left to weight different metrics on a scale from 0 to 10. In the background (on a seperate, hidden sheet), a rating will be calculated on a scale from 0 to 100 that includes the given weights. The country with the highest rating will be displayed with its overall rating and other KPIs which describe the metrics (also on a scale from 0 to 100). The KPIs change their color depending on the rating. Furthermore, the user may filter a specific continent and a diagram that is being displayed in the middle of the dashboard (map or bar chart).

<img src="Images/Template_Best_Country.png" width="500">

The **second dashboard** is rather on the technical side: It compares two different metrics, that the user is interested in. Several statistical variables that are calculated using all countries data are shown below the metrics. Two boxplots provide a concise overview of the distributions. On the right-hand side, there is a scatter plot to visualize the correlation between the selected metrics and a correlation matrix to show the correlations of all metrics (independent of the selection). Key insights at the bottom explain the selected metrics and the strength of the correlation for non-nerds.

<img src="Images/Template_Correlations.png" width="500">

## Data collection and cleansing

For data collection, you can use Excels feature to get data from the web: `Data -> Get Data -> From Other Sources -> From Web`. After adding an URL, Excels Power Query engine will extract the tables from the website by searching for HTML-tags such as `<table>`. You may now select the table you want to extract and save it locally as an Exce-file (in my repository the data is saved in the folder `'Separate_Data'`).

To merge the tables from sepatate tables/files into a "mother-table" that contains all data you can use the Power Query editor. First, you have to create connections for every file to later edit and transform them using Power Query: `Data -> Get Data -> From File -> From Excel Workbook -> Import -> Load to... -> Only Create Connection -> OK`. By double clicking on one of the queries in the right column, you will open the Power Query editor. Now you can merge the queries by navigating to `Home -> Merge Queries`. It makes sense to start from the the dataset with the most entries (in our case the countries by continent) to set a standard for the common column and successively merge the smaller datasets into the main one. You can now select the second table to merge with and the column, both tables share (the country-column). To not lose any data, you should select 'Full Outer' as a join type, as it keeps all rows from both tables.

<img src="Images/Join_Types.png" width="500">

I'd also not recommend the fuzzy matching, to stay in full control and get a better understanding of how to manually clean data.

To handle the entries that didn't match, you can replace the names of the countries in the newly merged table with the ones of the main table. Every change/transformation generates a new step in the column on the right and can be edited there or even be coded manually, if you are familiar with the M language. Although you could do this in the main table, if you perform the changes in the newly merged, separate tables, you'll keep the main table clean. In our case, the coutries that are often refered to differently are as follows:

* Czechia - Czech Republic - Czech Rep.
* DR Congo - Democratic Republic of the Congo - Democratic Republic of Congo
* Cape Verde - Capo Verde
* Turkey - Türkiye
* Unied Kingdom - Great Britain - Britain
* United Arab Emirates - UAE
* ...

Before loading the data, you may also delete unnecessary or double columns, change columns or their names, name transformations, set correct data types, and much more. You should never perfom any changes in the table you load into a spreadsheet! If you want to make changes using the Power Query editor later, the changes made in the sheet will be overwritten and won't apply to your data anymore.

## First Dashboard (Best Country)

<img src="Images/Screenshot_Best_Country.png" width="1000">

To get started, add two sheets to your workbook containing the data: One for the dashboad (`Best Country`), another for the calculations in the backgrond (`Best Country (help)`). In the help-sheet, you can now insert the data for this dashboard using dynamic arrays: `=IF(main_table[HDI (2023)]<>0; main_table[HDI (2023)]; "")`. The if-condition ensures that all empty entries which may contian a '0' will be cleared entirely. This is important, because we don't want cells containing '0' , as they would skew the calculations. Now, add new columns with standarsized values from 0 to 1: `=IFERROR((L5-MIN(L5#))/(MAX(L5#)-MIN(L5#)); "")`. You can also save the min- and max-values in separate cells to keep it tidy. If it is bad for the metric to have an high value (e. g. Gini coefficient), add `1-` before the calculation, because we want the standarsized values to go from 0 (bad) to 1 (good). The rating for each country is calculated as a weighted average of all standarsized metrics times 100:
$$\frac{\text{metric}_1\cdot\text{weight}_1 + ... + \text{metric}_n\cdot\text{weight}_n}{\text{weight}_1 + ... + \text{weight}_n}\cdot 100$$
The user can input the weights via spin buttons. To insert a spin button, you must first unhide the developer tab. Now you can insert form controls like spin buttons and connect them to a cell. You may also use scroll bars, but they don't tend to run as smoothly as spin buttons. In general, those form controls are very old Excel features and are not always integrated as smoothly as others. The data bars next to the spin buttons are conditionally formatted cells, depending on the value of the spin button.

Next, to only allow the use to enter certain values into the continent and diagram type selection, you can use data validation with a list of possible options on the help-sheet. Now filter the contries on the selected continent like so: `=FILTER(H5#; ((A34="World") + (A34=G5#)) * (I5:I238<>""))`. In my case, `A34` is the selected continent, `G5#` are the continents, `H5#` the countries and `I5:I238` the ratings. In calculations like this, Excel treats the value `TRUE` as 1 and `FALSE` as 0. `<>` is the negation-operator. Filter the ratings analogously. You can now link the country with the highest rating and the rating itself to your dashboard. To add KPIs (standarsized metrics) use `XLOOKUP()` to locate them and `ROUND()` to display as many digits as you please.

For the diagrams, I'd recommend to create a new sheet to keep things in perspective. Create both diagrams (or more) on the new sheet and style them as you like. Now copy the range of cells of one of your diagrams and paste it on the dashboard as a linked picture: `Right Click -> Paste Special -> Linked Picture`. Use the name manager to create a new name with the following condition: `=IF('Best Country'!$L$3="Map"; 'Best Country (diagrams)'!$A$1:$G$19; 'Best Country (diagrams)'!$A$22:$G$40)`. You now only have to insert the name you just set for the condition into the linked pictures formular bar and you first dashboard is finished.

## Second Dashboard (Correlations)

<img src="Images/Screenshot_Correlations.png" width="1000">
