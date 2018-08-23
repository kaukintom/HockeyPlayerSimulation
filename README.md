
## Fantasy Hockey Simulator
Simulates the point production of hockey players for an NHL season and accommodates a fantasy hockey scoring system. This project would be used to aid in fantasy hockey player selection.

Note: Created using VBA. However, files uploaded to Github containing the macros are classified as VB.

## Updates
Added the Excel workbook and VBA code which web scrapes wikipedia for birthdays and collects hockey player stats from Sportsnet.

Workbook name: "Hockey Data Collector.xlsm"

Note: Code for web scraping Sportsnet may need to be adjusted if there is a changed in the source code for the website

### Getting Started
Upon opening the Excel file "Hockey Player Point Projections.xlsm," there is a button that provides a tutorial on how to use the simulator.

1) The user will specify the amount of fantasy points recieved for all the relevent and available categories, then click the button that reads "Calculate Fantasy Points." Any category not containing a value, will not be included in the generated table.

2) To conduct simulation, then click the button that reads "Simulate Next Season's Point values." The user can specify how many simulations to conduct. The generated table from the previous step will include new columns of data, all related to the simulated results.

3) The user has the option to simulate the 2015/2016 season and have it compared to the 2016/2017 season. However, for convenience it is already simulated.

At any moment in time, as long as a table of hockey players is present on the worksheet (which there always should be), any relevent category to the table can be sorted by clicking button that reads "Sort By Category."

This workbook uses userforms for the tutorial, sorting, and simulation components. This helps guide the user and add convenience to the process.

### About The Simulation
Notes: 
- Data collection occured in another workbook, but the results can be viewed from the workbook "AdjustedStats - Latest Projections - Copy.xlsm" 
- All past goal and assist values are adjusted to the 2016/2017 level of goals per a game and assists per a game to accurately show when a player increases or decreases season to season 

The macros which are used to generate distributions are in the workbook "AdjustedStats - Latest Projections - Copy.xlsm," however, this workbook is cluttered, and the macros are not commented with detail. It will be difficult to understand, but it is available to be viewed. Also, the macros will not be able to run for this workbook since some of the sub procedures require connections to other workbooks. This workbook was recreated without the latest hockey season to produce the distributions for the previous year, to be able to test the error and predictability of the simulator. The second workbook is not placed in this repository, as it is completely identical to "AdjustedStats - Latest Projections - Copy.xlsm" but without the latest season of data.

Simulated results are based on how a player performed in the prior season. If it was a player's first season in the NHL, then it is based on how they are categorized (when first season occured and how they performed in their first season). The similated results are currently only goals and assists, which assume normality with the distributions.

Some category of players currently contain too few entries to develop an accurate probability distribution for, so some hockey players may not have their next season point totals simulated.

### Acknowledgments
McMaster Commerce course 3QA3 - Introduced simulation and how simulation can be achieved
Helped with collecting data from a table on a webpage:
https://stackoverflow.com/questions/34703533/how-to-scrape-data-from-the-following-table-format-vba

Helped with selecting options on a webpage from a dropdown menu:
https://www.ozgrid.com/forum/forum/other-software-applications/excel-and-web-browsers-help/145319-select-an-item-from-a-dropdown-list-on-webpage


