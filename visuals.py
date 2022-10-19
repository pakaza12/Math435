import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
import xlrd

sales_data = pd.read_excel("MitC2022data - SalesPopulation.xlsx")
valid_sales = sales_data[sales_data["SaleValidity"] == 1]
print(len(valid_sales["SaleValidity"]))


