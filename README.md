# Coffee Sales Excel Project
What I did:
1. Orders Table
   - Filled in columns (Customer Name, Email, Country, Coffee Type, Roast Type, Size, Unit Price, Sales, Coffee Type Name, Roast Type Name, Loyalty Card) by pulling data from Customers and Products tables mainly using XLOOKUP, INDEX + MATCH, and IF formulas.
   - Formulas used in each columns:
     - Customer Name
       =XLOOKUP(C2, customers!$A$1:$A$1001, customers!$B$1:$B$1001,,0)
     - Email
       =IF(XLOOKUP(C2, customers!$A$1:$A$1001, customers!$C$1:$C$1001,,0)=0,"",XLOOKUP(C2, customers!$A$1:$A$1001, customers!$C$1:$C$1001,,0))
     - Country
       =XLOOKUP(C2, customers!$A$1:$A$1001, customers!$G$1:$G$1001,,0)
     - Coffee Type
       =INDEX(products!$A$1:$G$49, MATCH(orders!$D2, products!$A$1:$A$49,0), MATCH(orders!I$1, products!$A$1:$G$1,0))
     - Roast Type
       =INDEX(products!$A$1:$G$49, MATCH(orders!$D2, products!$A$1:$A$49,0), MATCH(orders!J$1, products!$A$1:$G$1,0))
     - Size
       =INDEX(products!$A$1:$G$49, MATCH(orders!$D2, products!$A$1:$A$49,0), MATCH(orders!K$1, products!$A$1:$G$1,0))
     - Unit Prize
       =INDEX(products!$A$1:$G$49, MATCH(orders!$D2, products!$A$1:$A$49,0), MATCH(orders!L$1, products!$A$1:$G$1,0))
     - Coffee Type Name
       =IF(I2="Rob", "Robusta", IF(I2="Exc", "Excelsa", IF(I2="Ara", "Arabica", IF(I2="Lib", "Liberica"))))
     - Roast Type Name
       =IF(J2="M", "Medium", IF(J2="L", "Light", IF(J2="D", "Dark")))
     - Loyalty Card
       =XLOOKUP([@[Customer ID]], customers!$A$1:$A$1001, customers!$I$1:$I$1001,,0)

2. Pivot Tables & Charts
   - Total Sales: Showed the total monthly sales for each coffee type from the year 2019 to 2022
   - Top 5 Customers: Showed the top 5 customers with the highest number of sales
   - Sales by Country: Showed the total sales for each country

3. Dashboard
   - To visualize the total sales over time, top 5 customers, and sales by country.
   - Added timeline so users can see the data for the selected period.
   - Added slicers (coffee type, loyalty card, size) so users can filter data easily.
