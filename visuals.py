
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
import xlrd
import scipy

sales_data = pd.read_excel("MitC2022data - SalesPopulation.xlsx")
valid_sales = sales_data[sales_data["SaleValidity"] == 1]
#print(len(valid_sales["SaleValidity"]))

#plt.scatter(x=valid_sales["TLA"], y=valid_sales["Price"], title="Price vs. Total Living Area")
#plt.show()

fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(nrows=2, ncols=2)
ax1.plot(valid_sales["TLA"], valid_sales["Price"])
ax2.plot(valid_sales['Land Value', valid_sales['Price']])
ax3.plot(valid_sales['YrRD'], valid_sales['Price'])
ax4.plot(valid_sales['FixCt'], valid_sales['Price'])





#valid_sales_numeric = valid_sales[["CDU", "Qual", "TLA", "YrBlt", "LandValue", "Price"]]
#sns.pairplot(valid_sales_numeric)
#plt.show()


valid_price = valid_sales["Price"]
""" plt.hist(valid_price, bins=50, edgecolor='black')
plt.title('Histogram of Price')
plt.xlabel('Price')
plt.ylabel('Amount of Houses')
plt.show() """

""" price_mean = np.mean(valid_price)
print(price_mean)
xs = np.arange(0, 35113, 1)
price_sd= np.std(valid_price)
plt.plot(xs, scipy.stats.norm.pdf(valid_price, price_mean, price_sd))
plt.show()
 """

