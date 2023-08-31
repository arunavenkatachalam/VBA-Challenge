# VBA Challenge
 
# GOAL

To create a macro that would run on all the worksheets in the workbook "Multiple_year_stock_data". This macro would fetch and display the Ticker symbol, calculate the yearly change, percentage change and total volume of the particular ticker.

# Instructions

1. Intialized the variables to store the values and use it in the later part of the code. 

2. loop through all the worksheet to complete the task in the entire workbook 

3. Analyzed the worksheet to identify how the calculation can be performed and what is the difference in each sheet. Observed in the sheet "2018" the ticker value "AAB" ends in row number 252, whereas in the sheet "2019" first ticker value ends in row number 253 and in sheet "2020" the first ticker ends in 254. In order to incorporate this observation i have created an if condition to check the worksheet name and intialized a variable "r" and set the value to 250, based on the condition this value is changed for the second and third sheet.

4. Calculated the Last Row using the formula. Display the row header as mentioned in the Module 2 challenge.

5. Intialize the variable to store the Yearly_Change, Volume, Opening_Price, Closing_Price, Percentage_Change as data type Double to store the values respectively. Intialize row as Integer.

6. Intialize the object ColumnRange and TotalVolume as Range to store the entire column Percent Change and Total Stock Volume respectively and later it is used to calculate the maximum and minimum value in that range.

7. Loop the value starting from 2 to lastrow to perform the calculation to determine the ticker symbol, yearly change, Percent change and total stock volume based on the condition.

8. Calculate and display the highest increase and lowest decrease in Percent Change, highest value in total Volume along with the respective ticker symbol.

9. After performing the above steps for all the worksheets a popup window is displayed to inform the task is completed.
