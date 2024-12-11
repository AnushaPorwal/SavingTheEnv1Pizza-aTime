# Saving the environment one pizza at a time

## Introduction
Creates a visualization that shows the impact that production of certain crops have on the environment (in terms of the amount of water used and emissions released into the environment). It compares the data for 3 countries. It also displays the amount of these crops produced by these countries every year (between 1979 and 2019) as well as how much of that food was wasted.
The crops picked for this visualization are Wheat, Tomatoes, Peppers and Onions.
The countries picked are India, USA and Egypt.

The data for this project was obtained from:
1. Hannah Ritchie, Pablo Rosado and Max Roser (2022) - “Environmental Impacts of Food Production” Published online at OurWorldinData.org. Retrieved from: 'https://ourworldindata.org/environmental-impacts-of-food' [Online Resource]
2. Hannah Ritchie, Pablo Rosado and Max Roser (2022) - “Crop Yields” Published online at OurWorldinData.org. Retrieved from: 'https://ourworldindata.org/crop-yields' [Online Resource]
3. https://www.fao.org/platform-food-loss-waste/flw-data/en/


## Folder/File Guide
data/ - contains the cleaned data that was used for the visualization
dataCleaning/ - jupyter notebooks to clean the raw data
visualizationExcelMacros/ - contains the Excel Macros file that creates the individual charts to make up the final visualization
PPTcreationMacros/ - contains the ppt macros file, along with the separate module file
Saving the Environment one Pizza at a time.pdf - Project Report
Saving the Environment one Pizza at a time.mp4 - Final Visualization


## How to use or reproduce the code:
The Excel workbook contains the data, the different plots and the VBA scripts.
The sheet 'finalData2use - Copy' contains the data.
The 'Work' sheet contains pivot tables for the main data.
The 'Temp' sheet is for the code to be able to save the final visualization as an image. Do not delete it.


#### Description of what's found on 'Work' sheet
The year goes from 1979 to 2019, and the data is filtered year-wise.
We obtain total emissions and total water for the 3 countries for a given year in the 1st pivot table.
The 2nd pivot table contains the % of food wasted by each country in the given year.
The 3rd pivot table contains the amount of food produced by each country in the given year.
Finally there is a slicer to change the year from 1979 to 2019 one by one. (changing of year is done by the code)


#### Before running the VBA code to generate the plots, 3 things must be done:
1) Create the pie charts for the 3 countries, based on how much food is wasted. name the charts "cht\<Country>" Eg chtEgypt or chtUSA. You can also replace the smaller pie with an image of the pizza if you like.
2) Create a scatter plot that maps the country as a point on a graph of total emissions vs total water. name the chart "chtMain"
3) Make sure you create folders of name TempChartFolder and ResultCharts or change the folder paths in the code.

Once this is done, you should be able to run the code.

#### Code Description :: visualization-macros Excel file (folder visualizationExcelMacros)
**Module1.bas** contains the code to generate the different pie charts for each country and each year. The size of the pie chart is based on the amount of food produced by that country. The size of the pizza slice in the pie chart is based on the percentage of food wasted.
Run **subSavePizzas** to save all pie charts to a folder.

**Module2.bas** contains the code to generate the main chart for each year and save it to a new folder. It imports the individual pie charts for each year (obtained from running module 1) into the main chart. It sets the slider that depicts the year of the main chart. Once this is created, it saves this image. Run **subMakeMainChart** to save all main charts t

Both modules can be run separately - helps if you only want to regenerate/debug/test some part of the code.

**modGlobal.bas** has code to run both modules in order. Run **subMain** to run all the code together. Generates pie charts and the finals charts as well.
In this file you can set the image range of the pie charts. Min and max sizes are set to 2 cm and 7cm respectively. The code scales the min and max amounts produced to pie charts between 2cm and 7cm.

#### Code Description :: ToMakeTimeSpannedVideo PPT file (folder PPTcreationMacros)
For automatically generating the PPT that has one final plot per slide, run the VBA code associated with the PowerPoint Presentation file. It imports a final plot into each slide in order of year.
Run **ImportABunch** from **module1_PPT.bas** to create the PPT.

Once this is done, you can Export the PPT as a video (mp4) and set the time to be spent on each slide.
