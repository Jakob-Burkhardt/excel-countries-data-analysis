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

## Data scraping and cleaning