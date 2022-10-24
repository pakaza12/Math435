
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
import xlrd
import scipy

sales_data = pd.read_excel("MitC2022data - SalesPopulation.xlsx")
valid_sales = sales_data[sales_data["SaleValidity"] == 1]

plt.scatter(x=valid_sales["LandValue"], y=valid_sales["Price"], color="purple")
plt.title("Price vs. Land Value")
plt.ylabel("Price (in millions)")
plt.xlabel("Land Value")
plt.show()

valid_price = valid_sales["Price"]
