import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load data with absolute paths
q1_sales = pd.read_excel('C:/Users/Alfred Stephanus Dax/IdeaProjects/untitled/ProjectQ1Sales.xlsx')
q2_sales = pd.read_excel('C:/Users/Alfred Stephanus Dax/IdeaProjects/untitled/ProjectQ2Sales.xlsx')
q3_sales = pd.read_excel('C:/Users/Alfred Stephanus Dax/IdeaProjects/untitled/ProjectQ3Sales.xlsx')
q4_sales = pd.read_excel('C:/Users/Alfred Stephanus Dax/IdeaProjects/untitled/ProjectQ4Sales.xlsx')
products = pd.read_csv('C:/Users/Alfred Stephanus Dax/IdeaProjects/untitled/Project_products.csv')
standard_costs = pd.read_csv('C:/Users/Alfred Stephanus Dax/IdeaProjects/untitled/Project_standard_costs.csv')

# Strip whitespace from column names
q1_sales.columns = q1_sales.columns.str.strip()
q2_sales.columns = q2_sales.columns.str.strip()
q3_sales.columns = q3_sales.columns.str.strip()
q4_sales.columns = q4_sales.columns.str.strip()
products.columns = products.columns.str.strip()
standard_costs.columns = standard_costs.columns.str.strip()

# Print the first few rows of each DataFrame
print("Q1 Sales Data:")
print(q1_sales.head())
("\nQ2 Sales Data:")
print(q2_sales.head())
print("\nQ3 Sales Data:")
print(q3_sales.head())
print("\nQ4 Sales Data:")
print(q4_sales.head())
print("\nProducts Data:")
print(products.head())
print("\nStandard Costs Data:")
print(standard_costs.head())

# Merge dataframes
sales_data = pd.concat([q1_sales, q2_sales, q3_sales, q4_sales])
sales_data = sales_data.merge(products, on='Product_ID')
sales_data = sales_data.merge(standard_costs, on='Product_ID')

# Print the first few rows of the merged DataFrame
print("\nMerged Sales Data:")
print(sales_data.head())

# Check the columns in the merged DataFrame
print("\nColumns in Merged Sales Data:")
print(sales_data.columns)

# Check data types in the merged DataFrame
print("\nData Types in Merged Sales Data:")
print(sales_data.dtypes)

# Calculate metrics
if 'SalesAmount' in sales_data.columns and 'UnitCost' in sales_data.columns and 'QuantitySold' in sales_data.columns:
    sales_data['GrossProfit'] = sales_data['SalesAmount'] - (sales_data['UnitCost'] * sales_data['QuantitySold'])
    sales_data['ProfitPerUnit'] = sales_data['GrossProfit'] / sales_data['QuantitySold']
    sales_data['MarginPerUnit'] = (sales_data['ProfitPerUnit'] / sales_data['UnitPrice']) * 100

    # Analyze profitability by sales channel
    channel_profitability = sales_data.groupby('SalesChannel')['GrossProfit'].sum().sort_values(ascending=False)

    # Analyze profitability by product category
    category_profitability = sales_data.groupby('ProductCategory')['GrossProfit'].sum().sort_values(ascending=False)

    # Visualize profitability by sales channel
    plt.figure(figsize=(10, 6))
    sns.barplot(x=channel_profitability.index, y=channel_profitability.values)
    plt.title('Profitability by Sales Channel')
    plt.xlabel('Sales Channel')
    plt.ylabel('Gross Profit')
    plt.show()

    # Visualize profitability by product category
    plt.figure(figsize=(10, 6))
    sns.barplot(x=category_profitability.index, y=category_profitability.values)
    plt.title('Profitability by Product Category')
    plt.xlabel('Product Category')
    plt.ylabel('Gross Profit')
    plt.show()
else:
    print("One or more required columns are missing from the sales data.")