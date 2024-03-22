# Excel Dashboard - Coffee Sales

This is a Microsoft Excel project in which I created an interactive and dynamic dashboard showing interesting KPI's using mock coffee sales data.

This project helped me improve my skills in MS Excel using functions such as XLOOKUP, INDEX and IF statements. I also practiced creating pivot tables and interesting data visualizations. Finally, I refined my dashboarding skills and created appropriate filters and slicers that provided useful insights.  

Here are the steps I followed during the creation of this project:


### Step 1
Opened Excel file and reviewed the data. We have three (3) separate worksheets each containing data for Orders, Customers and Products. Each worksheet contains a primary key (Order ID, Customer ID & Product ID).


### Step 2
Converted the data into a table. We could do this in a later step, but I find the data is easier to read and clearer in a table format. Also, this will allow us to create pivot tables. 


### Step 3
Used XLOOKUP function to create 3 new columns (Customer name, Email, Country) in the "Orders" table. This is done using data from the "Customers" table and then transferring this data to the orders table. Thus, we are basically joining data from the customers table to the orders table (we will then use this data for analysis). The formula is the following: 
   
=XLOOKUP(C2,customers!$A$1:$A$1001, customers!$B$1:$B$1001,,0). 
   
I then clicked on the bottom right of the cell with the XLOOKUP function in order to auto-populate the column. I conducted similar steps for the email and country column. However, for the email column, I wrapped the XLOOKUP function within a IF statement, since we getting the value 0 for customers with no email. I used the following IF statement: 
  
=IF(XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001,,0)=0,"",XLOOKUP(C2,customers!$A$1:$A$1001, customers!$C$1:$C$1001,,0)) 
 
This will replace all 0 values with a null value.


### Step 4
Used the INDEX and MATCH functions to populate the columns Coffee Type, Roast Type and Size. I could use the XLOOKUP function as I did before, but the INDEX and MATCH funtions will allow me to populate all three columns using only one formula. I use the data from the product table within the formula. The formula is the following: 
  
=INDEX(products!$A$1:$G$49, MATCH(orders!$D2,products!$A$1:$A$49,0), MATCH(orders!I$1,products!$A$1:$G$1,0))

Finally, I can auto-populate all three columns


### Step 5
Created a Sales column using the following simple formula: =L2*E2. This is multiplying the "Quantity" column with the "Unit Price" column. I then auto-populated the rest of the column. 


### Step 6
Used an IF function to created column "Coffe Type Name" which will have the full name of the coffee and not just the 3 letter abbrivations we see in the "Coffee Type" column. The following formula is used : 

=IF(I2="Rob","Robusta",IF(I2="Exc","Excelsa",IF(I2="Ara","Arabica",IF(I2="Lib","Liberica",""))))

A similar procces is used to created column "Roast Type Name" that will include the full name instead of only the first letter as we see in column "Roast Type". The following formula is used: 

=IF(J2="M","Medium",IF(J2="L","Light",IF(J2="D","Dark","")))


### Step 7
Changed the data format of the "Order Date" column. I replaced the DD/MM/YYYY format with the DD-MMM-YYYY format. For example, instead of 01/01/2019, the format is 01-Jan-2019. I also changed the format for coffee size. Instead of only showing a number, all the data will also include a decimal and the letters "kg" at the end. For example, instead of 1, the format is now 1.0. Finally, I changed the format in the "Sales" column and included a dollar sign.


### Step 8
Used the "Remove Duplicates" feature to verify if the data had any duplicate rows. The data had no duplicates.


### Step 9
Inserted pivot table on a new worksheet using the data from the "Orders" table. I renamed the new worksheet "SalesLineChart". In this worksheet, I created a pivot table showing total sales for every coffee type by years and quarters. Next, I created a pivot chart using the data from the pivot table. This line chart will show total sales over time for every coffee type.


### Step 10
Created timeline using the pivot chart. This timeline will allow us to filter the data using dates. For example, selecting Q1-Q4 of 2019 in the timeline will make the pivot table (and, consequently, the pivot chart) only show data from 2019. 


### Step 11
Created the colunm "Loyalty Card" using XLOOKUP function. This is the formula: 

=XLOOKUP([@[Customer ID]],customers!$A$1:$A$1001,customers!$I$1:$I$1001,,0)


### Step 12
Inserted 3 slicers for "Size", "Roast Type Name" and "Loyalty Card". Theses slicers will also be used to filter the data in the dashboard. 


### Step 13
Copied the "TotalSales" worksheet and create a new worksheet which I renamed "CountryBarChart". Modified the pivot table by removing everything except "Sum of Sales" and added "Country". I then created a pivot chart using the data from the pivot table. This bar chart will show total sales for every country. 


### Step 14
Copied the "TotalSales" worksheet and create a new worksheet which I renamed "TopCustomers". Modified the pivot table by removing "Country" and added "Customer Name". Also, with the help of value filters feature, the pivot only shows the top 5 customers (customers with most sales). I then created a pivot chart using the data from the pivot table.


### Step 15
Copied the "TotalSales" worksheet and create a new worksheet which I renamed "SalesBarChart". Modified the pivot table by removing "Quarters". I then created a pivot chart using the data from the pivot table. This side-by-side bar chart will show total sales for every coffee type by year. 


### Step 16
Created two more pivot tables within a new sheet called "Dashboard". These pivot tables act like cards in Power BI and Tableau. The first pivot table shows the total sales and the second pivot table shows the total transactions. 


### Step 17
Added the timeline, all three slicers and every chart to the dashboard worksheet and verified that all filters are working properly.
