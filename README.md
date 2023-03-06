# VBA-challenge

Begining of Subscript
------------------------
The Sub-script presented here cycles through all three sheets to get the required data

The First secrion of code creates variable that will be used for the first data collection phase

first_value = stores the First day Open value of a market (EX: AAB)
last_value = stores the Closing value of the specific market (Ex:AAB)

first_found is a boolean so only the first value isn't overwritten in the loop


yearly_change will be used to store the change in value from last_value to first_value

stock_volume will store the total volume for each market (EX: AAB)

yearly chang
count is a varialbe used to help place the calculated values in the appropriate cells

The entire subscript is in a For Loop that cyles through each Sheet once it has been filed out with the required data


First For Loop
---------------
The first embedded For loop has two main Functions:
  It will cycle through the rows in column "A" checking that the current cell and the following one are of the same market (EX: AAB)
  On the first row it will save the first_value as the opening value of the market then turn first_found to true
  
  As the program goes through each row of the same market it will sum the total market volume 
  Once the next row is a different market the code will place the Total volume sum in it's respective cell, as well as the Ticker, and yearly change will be caclulated with the last_value-frist_value
  
 these actions will be repeated until the information for all Tickers has been completed 
 


The second For loop 
-------------------
This loop checks if the % change is negative or positive and colors it red or green respectively

Directly following this loop the row of %'s is formatted accordingly


THird For loop
-----------------
This loop cylces through the data collected in the first for loop and calculates greatest and least % change as well as greatest Volume
Following this loop the value are placed in their respective cells and labeled


Finally
---------

The subscript moves to the next sheeg and repeats the process
