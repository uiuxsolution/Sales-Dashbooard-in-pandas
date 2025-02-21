import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import calendar

# Load the data
@st.cache_data
def load_data():
    data = pd.read_excel('SampleData.xlsx', sheet_name='SalesOrders')
    data.drop_duplicates(inplace=True)
    data['OrderDate'] = pd.to_datetime(data['OrderDate'])
    data['OrderYear'] = data['OrderDate'].dt.year
    data['OrderMonth'] = data['OrderDate'].dt.month
    return data

data = load_data()

# Sidebar filters
st.sidebar.header('Filters')
year = st.sidebar.slider(
    'Select Year:',
    min_value=data['OrderYear'].min(),
    max_value=data['OrderYear'].max()
)
filtered_data = data[data['OrderYear'] == year]

# Main app
st.title('Sales Dashboard')
st.subheader('Year: {}'.format(year))

# Sales by region
st.subheader('Sales by Region')
fig1, ax1 = plt.subplots()
sns.barplot(x='Region', y='Total', data=filtered_data, estimator=sum, ax=ax1)
st.pyplot(fig1)

# Sales by representative
st.subheader('Sales by Representative')
fig2, ax2 = plt.subplots()
sns.barplot(x='Rep', y='Total', data=filtered_data, estimator=sum, ax=ax2)
plt.xticks(rotation=45)
st.pyplot(fig2)

# Monthly sales
st.subheader('Monthly Sales Trend')
fig3, ax3 = plt.subplots(figsize=(12, 6))
monthly_sales = filtered_data.groupby('OrderMonth')['Total'].sum().reset_index()
monthly_sales['Month'] = monthly_sales['OrderMonth'].apply(lambda x: calendar.month_abbr[x])
sns.barplot(x='Month', y='Total', data=monthly_sales, ax=ax3)
st.pyplot(fig3)

# Top selling items
st.subheader('Top Selling Items')
top_items = filtered_data.groupby('Item')['Units'].sum().reset_index().sort_values(by='Units', ascending=False).head(5)
st.table(top_items)

# Total sales vs Total revenue (assuming revenue is Units * Unit Cost)
st.subheader('Sales & Revenue Correlation')
fig4, ax4 = plt.subplots()
sns.scatterplot(data=filtered_data, x='Total', y='Units', ax=ax4)
st.pyplot(fig4)

# Summary statistics
st.subheader('Summary Statistics')
st.table(filtered_data.describe())