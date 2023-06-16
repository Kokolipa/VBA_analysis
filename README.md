# Stock Analysis with VBA
### Project Description
The project includes a VBA solution I built to analyse generated stock market data. The project contains two datasets. The sample file ('alphabetical_texting) with six worksheets, 20k+ records each, and a more extensive dataset ('Multiple_year_stock_data')  with three worksheets that have over 750k records per worksheet. 

#### Folder structure
----
``` yml
.
├── VBA_Assignment
│   ├── Images                          # This is where I stored the spreadsheets result images
│   │   ├── first_sheet.png    
│   │   ├── second_sheet.png 
│   │   ├── third_sheet.png
│   └── ..                  
|   ├── Resources                       # This folder contains the xlxm fils
│   |   ├── alphabetical_testing.xlsx.xlsm            
│   |   ├── Multiple_year_stock_data.xlsm                    
│   └── ...       
|   ├── Stock_Classifier.vbs            # This is the VBA script             
|              
|___README.md
``` 

### About the VBA Solution
The VBA script iterates through each worksheet within a given workbook and provides as an output the:
1. **Stock ticker**. 
2. The **yearly difference** between the stock's closing price and the opening price. 
3. The **percentage difference** relative to the opening price. 
4. The **total stock volume**. 

In addition, the script creates a summary table that includes the following: 
1. Greatest percentage increase and decrease from the **percentage difference column**.  
2. Greatest total volume from the **total stock volume column**. 

![first_sheet](https://github.com/Kokolipa/VBA_analysis/blob/stocks/VBA_Assignment/Images/first_sheet.png)
