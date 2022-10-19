import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
import xlrd

sales_data = pd.read_excel("MitC2022data - SalesPopulation.xlsx")
valid_sales = sales_data[sales_data["SaleValidity"] == 1]
#print(len(valid_sales["SaleValidity"]))

#plt.scatter(x=valid_sales["TLA"], y=valid_sales["Price"], title="Price vs. Total Living Area")
#plt.show()

valid_sales_numeric = valid_sales[["CDU", "Qual", "TLA", "YrBlt", "LandValue", "Price"]]
sns.pairplot(valid_sales_numeric)
plt.show()

