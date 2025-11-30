import pandas as pd
import numpy as np

np.random.seed(42)

n = 10000  # num of rows

# Sales Transactions
branches = ['Nasr City', 'Maadi', 'Heliopolis', 'Dokki', '6th October']
products = ['Smartphone A', 'Smartphone B', 'Smartphone C', 'Smartwatch', 'Earbuds', 'Charger', 'Powerbank']
customer_type = ['Regular', 'New', 'VIP']
dates = pd.date_range(start='2025-01-01', end='2025-11-30', freq='H')
dates = np.random.choice(dates, n)
sales_df = pd.DataFrame({
    'Sale ID': range(1, n+1),
    'Date': dates,
    'Branch': np.random.choice(branches, n),
    'Product': np.random.choice(products, n),
    'Quantity': np.random.randint(1, 5, n),
    'Unit Price': np.random.randint(1000, 15000, n),
    'Customer ID': np.random.randint(1000, 2000, n)
})
sales_df['Total Sales'] = sales_df['Quantity'] * sales_df['Unit Price']

# Products Info
categories = ['Smartphones', 'Accessories', 'Wearables']
suppliers = ['Supplier A', 'Supplier B', 'Supplier C']
products_df = pd.DataFrame({
    'Product': products,
    'Category': np.random.choice(categories, len(products)),
    'Cost': np.random.randint(500, 10000, len(products)),
    'Supplier': np.random.choice(suppliers, len(products)),
    'Warranty (Months)': np.random.randint(6, 25, len(products))
})

# Branches Info
managers = ['Manager A', 'Manager B', 'Manager C', 'Manager D', 'Manager E']
regions = ['East', 'South', 'North', 'West', 'Central']
branches_df = pd.DataFrame({
    'Branch': branches,
    'Manager': managers,
    'Region': regions,
    'Opening Date': pd.to_datetime(['2018-01-15','2019-03-20','2017-07-10','2016-05-05','2020-09-01'])
})

# Customers Info
customer_ids = range(1000, 2000)
names = [f'Customer_{i}' for i in customer_ids]
join_dates = pd.date_range(start='2020-01-01', end='2025-11-01')
customers_df = pd.DataFrame({
    'Customer ID': customer_ids,
    'Customer Name': names,
    'Type': np.random.choice(customer_type, len(customer_ids)),
    'Join Date': np.random.choice(join_dates, len(customer_ids)),
    'Loyalty Points': np.random.randint(0, 5000, len(customer_ids))
})

# save
file_path = 'Sales_Training.xlsx'
with pd.ExcelWriter(file_path) as writer:
    sales_df.to_excel(writer, sheet_name='Sales_Transactions', index=False)
    products_df.to_excel(writer, sheet_name='Products_Info', index=False)
    branches_df.to_excel(writer, sheet_name='Branches_Info', index=False)
    customers_df.to_excel(writer, sheet_name='Customers_Info', index=False)

print(f'Excel file created: {file_path}')
