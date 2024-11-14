# Excel-RMS-Sales-Analysis

This project analyses customer orders for RMS to provide key insights into sales performance, customer order behaviour, and operational efficiency. The goal is to enhance decision-making by examining trends in product profitability, order fulfillment times, and customer purchasing frequency.
---

## Data Preparation

- **Formatting**:
    - Formatted the *Sales* and *Profit* columns to two decimal places and added thousand separators for readability.
    - Changed data types for *Product Base Margin*, *Shipping Cost*, *Unit Price*, and *Profit* columns from *General* to *Number*.
- **Handling Missing Values**:
    - Replaced blanks in the *Product Base Margin* column with "N/A" to maintain data integrity.
- **Date Formatting**:
    - Changed the *Order Date* column to the *Date* data type.
    - Added an *Order Year* column next to the *Order Date* using the formula `=YEAR(@OrderDate)` to enable year-based analysis.
- **Order Fulfillment Time**:
    - Created a *Days to Ship* column to calculate order fulfillment time. Used the formula `=DATEDIF(OrderDate, Shipdate, "D")` to measure the number of days from order placement to shipment.
- **Cost of Goods Sold (COGS)**:
    - Added a *COGS* column by calculating the total cost for each order. This was done by multiplying *Unit Price* by *Quantity* and adding the *Shipping Cost*. This allowed for accurate *Gross Profit* calculation.
- **Enhanced Returns Table**:
    - Used *INDEX MATCH* functions in the Returns table to add columns for *Customer Segment*, *Order Date*, and *Product Category*, providing more context to returned items.
- **Customer Purchase Behaviour**:
    - Added an *Order Frequency* column to determine how many times a customer has purchased from the business. This was part of a planned RFM (Recency, Frequency, Monetary) analysis to segment customers based on purchase behaviour.

---

## Analysis & Findings

The analysis section is based on key performance indicators calculated using pivot tables, which include insights on sales, profit, costs, and customer metrics.

- **Sales & Quantity**:
    - **Total Sales**: $14,915,601
    - **Total Order Quantity**: 214,777 units
    - **Average Order Value (AOV)**: $69.45
- **Profit & Costs**:
    - **Total Profit**: $1,521,768
    - **Cost of Goods Sold (COGS)**: $15,615,006
    - **Gross Profit Margin**: -5%
    - **Total Shipping Cost**: $107,831
- **Order Metrics**:
    - **Total Orders**: 8,399 orders
    - **Average Days to Ship**: 2 days
    - **Return Rate**: 6.8% (572 returned orders)

**Product Category Breakdown**

This analysis examines sales, profit contributions, and order counts across key product categories: **Technology**, **Furniture**, and **Office Supplies**.

- **Order Count by Category**:
    - **Office Supplies**: 4,610 orders
    - **Technology**: 2,065 orders
    - **Furniture**: 1,724 orders
- **Sales and Profit by Category**:
    - **Technology**: $5,984,248 in sales with $886,314 in profit
    - **Furniture**: $5,178,591 in sales with $117,433 in profit
    - **Office Supplies**: $3,752,762 in sales with $518,021 in profit

![Screenshot 2024-11-12 at 21 13 29](https://github.com/user-attachments/assets/ef4852c8-bb26-46dd-bf7c-cf17c5c02103)


- **Top 5 Products Ordered**:
    - **Paper**: 1,225 orders
    - **Binders and Binder Accessories**: 915 orders
    - **Telephones and Communication**: 883 orders
    - **Office Furnishings**: 788 orders
    - **Computer Peripherals**: 758 orders
- **Bottom 5 Products Ordered**:
    - **Rubber Bands**: 179 orders
    - **Scissors, Rulers and Trimmers**: 144 orders
    - **Copiers and Fax**: 87 orders

![Screenshot 2024-11-12 at 21 16 05](https://github.com/user-attachments/assets/db1fdcbc-3b1b-4936-b6b3-53ef8410c1bb)

**Customer Segment**

- **Corporate**:
    - Orders: 3,076
    - Contribution: Corporate customers place the highest number of orders, suggesting high engagement and repeat purchases.
- **Home Office**:
    - Orders: 2,032
    - Contribution: This segment also shows significant engagement, with over 2,000 orders.
- **Consumer**:
    - Orders: 1,649
    - Contribution: Consumer orders are moderate, which could indicate opportunities for increasing engagement.
- **Small Business**:
    - Orders: 1,642
    - Contribution: Small Businesses have a similar order volume to Consumers, indicating comparable engagement levels.

![Screenshot 2024-11-12 at 21 21 55](https://github.com/user-attachments/assets/dd3d064b-26c1-4e15-bf8b-b3b39e01b887)
![Screenshot 2024-11-12 at 21 27 26](https://github.com/user-attachments/assets/30426b03-84a6-4ed5-9073-0c4ddf355827)


- **Corporate and Home Office** segments are the top contributors to both sales and order volume, suggesting that these segments are highly engaged and profitable.
- **Small Business and Consumer** segments, while profitable, have lower order volumes, which may represent growth opportunities through targeted marketing efforts or customer loyalty programs.

### Returns Summary

- **Total Returns**: 572, with Corporate and Home Office segments having the highest return rates.
- **Returns by Year**: Slight decrease from 2009 to 2012, potentially indicating quality improvement.
- **Top Product Returns**:
    - **Binders and Binder Accessories** (77 returns) and **Office Furnishings** (64 returns) lead in returns, suggesting areas for quality checks.

This analysis highlights Corporate and Home Office segments, along with certain office supplies, as focal points for quality or service improvements to reduce returns.

![Screenshot 2024-11-12 at 21 41 26](https://github.com/user-attachments/assets/218d4ae9-80c8-4274-ab1a-4a826f97606e)


**Ship Mode** 

**Order Counts**:

- **Regular Air**: 6,270 orders
- **Delivery Truck**: 1,146 orders
- **Express Air**: 983 orders

**Total Shipping Costs**:

- **Regular Air**: $48,008
- **Delivery Truck**: $51,972
- **Express Air**: $7,851
- **Regional Breakdown**:
    - The **West** and **Ontario** regions have the highest number of orders, with **Regular Air** being the most common mode.
    - **Express Air** is used more in regions like the Atlantic and Prairie.

![Screenshot 2024-11-12 at 22 03 27](https://github.com/user-attachments/assets/f43bd94f-dbc0-4474-8b39-0dba1ed3c97d)

To identify the most valuable customers, I conducted an RFM (Recency, Frequency, Monetary) analysis. I focused on customers with purchases in 2012 for Recency, those with 15–20 purchases for Frequency, and the highest spenders for Monetary value. This analysis revealed the top 10 high-value customers. Using INDEX MATCH, I determined these customers' primary purchases, allowing for a detailed subcategory breakdown. The high-value customers mainly focus on Office Supplies and Furniture, particularly "Paper," "Binders and Binder Accessories," "Tables," and "Office Furnishings." This concentration suggests an opportunity for targeted marketing of office essentials to increase retention and drive sales. While less prominent, interest in Technology items like "Computer Peripherals" and "Telephones and Communication" highlights a potential niche market.

![Screenshot 2024-11-14 at 21 14 35](https://github.com/user-attachments/assets/d60182c3-7a6c-424b-b8e9-55b156c5b765)


### **Key Findings**

- The Gross Profit Margin over the 3 years of sales is consistently in the negative which poses profitability issues which suggests the company needs to either review product discounts or review product prices
- All the products in the top 5 most ordered products are also in the top 10 most returned products which would indicate that there’s a quality issue and this needs to be addressed by the company
- 'Critical' priority orders have an average fulfillment time of 2 days compared to 'high' and 'medium' priority orders at just 1 day, indicating a misalignment in prioritising important orders. The company needs to implement a system ensuring that critical orders receive top priority and are fulfilled first, while medium priority orders can take up to 2 days.
- About 75% of orders are shipped using regular air, which is neither the fastest nor the cheapest shipping mode. The company spends the most on delivery trucks—the most economical but slowest option—accounting for 12% of orders shipped. This indicates the company is managing costs well. However, there's concern that critical and high-priority orders rarely use express air. The company spends approximately $8,000 on express air compared to $50,000 on other shipping modes. Charts show roughly 1,600–1,700 orders for each shipping mode, suggesting the company should increase express air usage to ensure timely delivery of priority orders. This current approach could negatively impact customer retention, prompting a need to review the shipping structure.
- 
