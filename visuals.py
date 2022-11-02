
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
import xlrd
import scipy

sales_data = pd.read_excel("extraFilter.xlsx")

""" plt.scatter(x=valid_sales["LandValue"], y=valid_sales["Price"], color="purple")
plt.title("Price vs. Land Value")
plt.ylabel("Price (in millions)")
plt.xlabel("Land Value")
plt.show() """

valid_price = sales_data["Price"]
sd = np.std(valid_price)
mn = np.mean(valid_price)
# normal
a_mu, a_std = scipy.stats.norm.fit(valid_price)
xs = np.linspace(0, max(valid_price))
p = scipy.stats.norm.pdf(xs, a_mu, a_std)



sns.displot(data=sales_data, x='Price', kde=True, bins=50, color = "green")
plt.axvline(x=mn+(2*sd),color="steelblue", linestyle="dashed")
plt.axvline(x=0,color="steelblue", linestyle="dashed")

plt.xlabel('Price (in millions)')
plt.ylabel('Number of homes')
plt.title('Price Distribution')
plt.show()

""" plt.hist(valid_price, bins=30, color = "green", edgecolor="black")

plt.plot(xs, p)
#plt.xlim(0, 1.5e6)

plt.show() """






