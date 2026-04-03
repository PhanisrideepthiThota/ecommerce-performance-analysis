import pandas as pd

# File paths (same folder)
import pandas as pd

orders = pd.read_excel(
    "../data/Worksheet in E-Commerce _Command Center (1).xlsm",
    sheet_name="List of Orders"
)

order_details = pd.read_excel(
    "../data/Worksheet in E-Commerce _Command Center (2).xlsm"
)

sales_target = pd.read_excel(
    "../data/Worksheet in E-Commerce _Command Center (3).xlsm"
)
#verify the data is loaded correctly
print(orders.head())
print(order_details.head())
print(sales_target.head())

print(orders.columns)
print(order_details.columns)
print(sales_target.columns)

#convert oder date to datetime
orders['Order Date'] = pd.to_datetime(orders['Order Date'])

#create month column in orders
orders['Month'] = orders['Order Date'].dt.to_period('M').astype(str)

#converts sales target month column
sales_target['Month'] = pd.to_datetime(
    sales_target['Month of Order Date']
).dt.to_period('M').astype(str)

#merge dattasets
merged_df = pd.merge(
    order_details,
    orders,
    on='Order ID',
    how='inner'
)

#Monthly Profit
monthly_profit = merged_df.groupby('Month')['Profit'].sum().reset_index()

#visualise-Is the company growing or declining over time?
import matplotlib.pyplot as plt

plt.figure(figsize=(10,5))
plt.plot(monthly_profit['Month'], monthly_profit['Profit'], marker='o')
plt.xticks(rotation=45)
plt.title("Monthly Profit Trend")
plt.xlabel("Month")
plt.ylabel("Profit")
plt.tight_layout()
plt.show()

#PROFIT BY STATE
state_profit = (
    merged_df.groupby('State')['Profit']
    .sum()
    .sort_values(ascending=False)
)
#visualise-This shows regional performance.
plt.figure(figsize=(10,5))
state_profit.plot(kind='bar')
plt.title("Profit by State")
plt.xlabel("State")
plt.ylabel("Profit")
plt.tight_layout()
plt.show()

#CUSTOMER FOCUS ANALYSIS
top_customers = (
    merged_df.groupby('CustomerName')['Profit']
    .sum()
    .sort_values(ascending=False)
    .head(10)
)
#visualise-High-value customers should be retained.
plt.figure(figsize=(10,5))
top_customers.plot(kind='bar')
plt.title("Top 10 Profitable Customers")
plt.xlabel("Customer")
plt.ylabel("Profit")
plt.tight_layout()
plt.show()

#CATEGORY & SUB-CATEGORY CONTRIBUTION
category_profit = merged_df.groupby('Category')['Profit'].sum()

#Visualise-Identifies profitable vs risky product lines.
category_profit.plot(kind='bar')
plt.title("Profit by Category")
plt.xlabel("Category")
plt.ylabel("Profit")
plt.tight_layout()
plt.show()

#SALES TARGET ANALYSIS- Which categories met targets/months performed well

#Actual profit by Category & Month
actual_profit = (
    merged_df.groupby(['Category','Month'])['Profit']
    .sum()
    .reset_index()
)

#Merge with Sales Target
target_vs_actual = pd.merge(
    actual_profit,
    sales_target[['Category','Month','Target']],
    on=['Category','Month'],
    how='inner'
)

#Target achievement flag
target_vs_actual['Target Met'] = (
    target_vs_actual['Profit'] >= target_vs_actual['Target']
)

#CORRELATION MATRIX
import seaborn as sns

corr_matrix = merged_df[['Amount','Profit','Quantity']].corr()

plt.figure(figsize=(6,4))
sns.heatmap(corr_matrix, annot=True, cmap='coolwarm')
plt.title("Correlation Matrix")
plt.tight_layout()
plt.show()
