# DataScience_Assignment_2
January 28, 2019  
Scott McEachern

This repo contains the files for the Assignment 2;  "The VBA of Wall Street".  The assignment has three different files that are used to update the Excel document.  Outlined below are the steps to deploy.

### 1 Download Files
Download the following VBA files from the GitHub repo:  
* GreatestTickerInfo.cls
* TickerInfo.cls
* YearlyStockCalculator.bas

### 2 Open Excel
Open Excel document that calculations are to be preformed; such as Multiple_year_Stock_Data.xlsx

### 3 Visual Basic Editor
Within the newly opened Excel document, open the Visual Basic editor.

### 4 Load Files
Within the Visual Basic Editor, load in the previously downloaded files.  From the Project Explorer, 
right click the “VBA Project” node and from the pop-up menu select “Import” and then navigate to the files.  

After the three files have been imported, there will be one “Module” named “YearlyStockCalculator” 
and two “Class Modules” named “GreatestTicketInfo” and “TicketInfo”.  

![VBA Project Explorer](https://github.com/normalfactory/DataScience_Assignment_2/blob/master/Instructions/ProjectExplorer.png)

### 5 Run Macro
Open the Marco dialog and run the the macro named “Start”.  

A message box is displayed when the processing is completed.  The processing has been found to take 364 to 430 seconds.  

![Message box displayed when completed](https://github.com/normalfactory/DataScience_Assignment_2/blob/master/Instructions/CompletedMessage2.png)
