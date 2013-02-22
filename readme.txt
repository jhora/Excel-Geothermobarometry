Computer files related to "An Excel-based tool for evaluating and visualizing geothermobarometry data" by JM Hora, A Kronz, S Moeller-McNett, and G Woerner.

1. Purpose and overview

This file is intended as a user's manual for the Excel geothermometry workbooks. 

The steps involved are in principle straightforward:

(1) Paste your data into the yellow-filled cells.
(2) Drag formulas to match your data ranges
(3) Paste analyses numbers corresponding to mineral phases along top and left sides of the calculation grid.
(4) Drag formulas to match your data range.
(5) Temperatures are either automatically calculated, or in the case of Fe-Ti Oxides, the user opens a third-party program and activates a script, which act together to do the calculations and place the results into the Excel sheet automatically.

The remainder of this manual goes into detail about the above steps, should it be needed to troubleshoot problems.  It also describes the (minor) differences in spreadsheets for the various thermometers. Please note the sections marked as ‘important’, should you need to modify the spreadsheet by inserting cells.

General information about the utility of the approach as well as results of the test case study can be found in the main text of the article. Here, we describe how the user can import data into the templates and how to adapt the templates based on individual project requirements.

This manual assumes that the reader has access to the following files, available as an electronic appendix to the article.

Hbl-Plag.xlsx
Hbl-Plag_worked_example.xlsx
Two-Fsp.xlsx
Two-Fsp_worked_example.xlsx
Fe-Ti_Oxide.xlsx
Fe-Ti_Oxide_worked example.xlsx
Oxide_Data_Transfer.scpt

2. System Requirements

2.1 For Hbl-Plag and Two-Fsp workbooks

Excel 12.0 (MS Windows version 2007 or Apple Macintosh version 2008) or later

Earlier versions of Excel will be able to perform calculations, but will not display conditional formatting (color coding of results useful for visualization).

2.2 For Fe-Ti Oxide workbook and Oxide data transfer.scpt

Mac OSX and Excel version 12.0 (2008) or later are needed to run the AppleScript that facilitates the data transfer between Excel and the FeTiOxideGeotherm application (distributed by ofm-research.org).
We have found that the FeTiOxideGeotherm application does not function properly using OSX v.10.5 (Leopard), but runs as designed under v.10.6 (Snow Leopard), v.10.7 (Lion), and v.10.8 (Mountain Lion).

3. Instructions

3.1 General instructions for entering data

In this and the following sections, it may be helpful to open and refer to the appropriate worked example file.

For all geothermometry worksheets, yellow-filled cells do not contain formulas are intended for user input. Cells with other fill colors and those that contain formulas in general should not be altered.  None of the cells are locked, such that the user is free to copy and insert columns and rows in order to accommodate data sets of (nearly) arbitrary size. Each workbook consists of worksheets of two types: 

Tables: one or more worksheets into which mineral EPMA analyses can be copied and which performs mineral formula recalculation. These are labeled with mineral names.

Grids: one or more worksheets, whose purpose is to lookup data from Table Worksheets and use that data to perform geothermometry calculations, using the grid arrangement described in the accompanying article. These are labeled with names of thermometers.

In Table worksheets, the user must input (or copy) data according to the column headings. All formulas are written such that each row corresponds to a singe analysis; entire rows of formulas can be copied down to accommodate large data sets. In general, columns should NOT be inserted in the data entry worksheets, except for the purpose of calculating a parameter that will be later used for sorting.  

IMPORTANT: In no case should the user insert columns within the region labeled 'Lookup table' at the far right side of the data entry worksheet.  Doing this will cause the Lookup functions used in Grid worksheets to not function properly, and will give meaningless calculated temperatures.

In Grid worksheets, the user must enter (or copy-paste) only the analysis numbers in the appropriate yellow-filled cells. Lookup functions automatically populate the cells to the right (or below) each yellow highlighted analysis number with parameters appropriate for that analysis to be used as inputs in the geothermometry calculation. The cells containing the lookup functions can be copied down and across, respectively in order to accommodate data sets of any size (up to 500 analyses of each phase - for even larger data sets, see Sec. 4, below). Grid calculation formulas are also written such that they can be copied across and down, and (except for Fe-Ti Oxides, see below) calculate resulting temperatures. Conditional formatting will automatically fill the cell with a color indicative of temperature, with a color along a gradient where blue = 600°C, red = 900°C, and yellow = 1200°C.  The user is free to modify this by changing the conditional formatting rules applicable to those cells, for more information see appropriate section of the Excel help menu.

Grid worksheets are designed such that the oder of analyses can be sorted according to any parameter, that can be freely chosen for each mineral.  The worksheet does this by looking up this parameter from the section of Table worksheets labelled 'Lookup table'.  This is located to the right of all other formula recalculation columns. In order to set this sort parameter, the user must insert references to that data into the cells of the yellow highlighted column in the appropriate Table worksheets. Once this is done, the lookup function will automatically populate all Grid worksheets with this data. Once this has taken place, the calculation grid can be sorted as desired.

3.2 Specific instructions for Hbl-Plag

This workbook calculates mineral formulas according to 23 oxygens following the method from Holland and Blundy (1994). The user can change this oxygen basis in Cell AM3, but this is not recommended as the thermometer is predicated on using 23 oxygens. The Table worksheet for Amphibole also recalculates the mineral formula as shown in Ridolfi et al. (2010), and calculates P, T, fO2, and melt composition based on Ridolfi et al. (2010) and Ridolfi and Renzulli (2012).  The T and P results are transferred to the grid (two rows of color formatted cells with one result per amphibole analysis) for comparison with temperatures obtained from Holland and Blundy (1994) grid calculations. These temperatures and pressures are only applicable to calc-alkaline volcanic rocks, for other applications, the temperatures for these thermometers are not expected to be in close agreement.

Two Grid worksheets calculate temperatures based on the Edenite-Tremolite and Edenite-Richterite formulations, respectively. Owing to the polybaric crystallization histories of many volcanic rocks, the pressure input for the Holland and Blundy grid calculation is taken to be the amphibole pressure estimate from Ridolfi and Renzulli 2012. If the user wishes to use a constant P, the appropriate value can be simply pasted into these cells.
3.3 Specific instructions for Two-Fsp

This workbook calculates mineral formulas according to 8 oxygens. The user can change this oxygen basis in Cell AC3, but this has no effect on subsequent calculations, as cation ratios are used.

One Grid worksheet calculates temperatures based on the equation 27 of Putirka (2008). This requires the user to input a pressure estimate in cell C3.

3.4 Specific instructions for Fe-Ti Oxides

The Fe-Ti Oxides.xlsx workbook differs from the others, in that there is a single worksheet where all oxide analyses are pasted.  Because it is difficult to distinguish Fe-rich endmembers of the Hematite-Ilmenite and Magnetite-Ulvospinel solid solutions in the absence of information on the oxidation state of Fe, this worksheet makes this identification by calculating both mineral formulas for each analysis, and accepting the one that gives a total closer to 100%, while rejecting the other.  Disclaimer: this process works well for 'good' analyses where the process produces an unambiguous result, but will provide poor identification if there are significant amounts of un-analysed minor elements and/or if the analyses are poorly calibrated, leading to low or high totals. The worksheet allows for rejection of analyses where the best-fit total is either too low or too high, indicative of poor mineral identification. This threshold can be set by the user in cell BW5.

In order to enter the analysis numbers in the worksheet 'Ghiorso & Evans (2008)', the 'Data Entry' worksheet can be sorted according to the 'Mineral' column (column BY), and the appropriate analysis numbers copy-pasted.

Unike the other, stand-alone thermometry workbooks described in Sec. 3.2 and 3.3, Fe-Ti Oxide.xlsx requires a third-party application to do the thermometry calculations. Data transfer between Excel and this application is done by the provided Oxide Data Transfer.scpt.
Before running the script, be sure that the following conditions are met:
(1) In System Preferences > Universal Access, make sure that the box for 'Enable access for assistive devices' is checked.
(2) Excel is running and the file Fe-Ti Oxide.xlsx is the only Excel workbook open, and that the active worksheet is 'Ghiorso & Evans (2008)'. It is allowed to rename this file and worksheet, however, note that the script is written such that it will always use the worksheet that is open and active at the time of script launch.
(3) The application FeTiOxideGeotherm is open.

IMPORTANT: Do not insert or delete columns or rows in the upper left area of the 'Ghiorso & Evans (2008)' worksheet, and make sure that the upper left-most cell in the calculation grid is Q17. If this is not the case, the script will fail to initialize the grid calculation in the appropriate cell, and the script will return an error, or make calculations based on wrong inputs.

To run the script, open Oxide Data Transfer.scpt in AppleScript Editor and click run. AppleScript Editor allows the user to view the code of the script, make changes and recompile, and also can be used to stop the script partway through the run.

4. Miscellaneous notes and troubleshooting

Large data sets: For data sets including more than 500 analyses, the user must enlarge the lookup range (the second argument of the lookup function in order to include all cells within the lookup table). Make sure this is copied to all the cells that use the lookup function affected.

The VLOOKUP functions in these workbooks use FALSE is used as the last argument. This searches for only exact matches on sample number. Consequently, if there is an error with stoichiometery that results in a blank cell in the lookup table, and an exact match cannot be found, the #N/A error will be returned in the calculation inputs area of the grid.

5. Adapting the template for use with another thermometer

The workbooks provided are unlocked, and therefore it is possible for the user to view all formulas and to make modifications to suit their individual needs. For more information on syntax of the VLOOKUP functions that are extensively used here, or how to modify conditional formatting, please see the appropriate MS Excel help pages.
