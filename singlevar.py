import numpy as np
import pandas as pd
from scipy.stats import t
import math
sales_data = pd.read_excel("MitC2022data - SalesPopulation.xlsx")
valid_sales = sales_data[sales_data["SaleValidity"] <= 2]

def single_var():
    col_name = input('What column would you like to analyze? ')

    current_var = valid_sales[col_name]

    #current_var = current_var.loc[(current_var != 0).any(axis=1)] removing 0s?

    current_var.dropna(axis=0)

    # mean and sd
    mean = np.mean(current_var)
    sd = np.std(current_var)

    # confidence interval
    dof = len(current_var) - 1
    alpha = 0.95
    stan_err = sd/math.sqrt(len(current_var))
    t_crit = np.abs(t.ppf((alpha)/2,dof))

    lower_ci = mean - (t_crit*stan_err)
    upper_ci = mean + (t_crit*stan_err)

    confidence = (round(lower_ci,2), round(upper_ci,2))

    print('mean:', round(mean, 2), "\nstandard deviation:", round(sd, 3))
    print('95% confidence of mean:', confidence)

single_var()