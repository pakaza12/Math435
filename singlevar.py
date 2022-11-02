from statistics import median
import numpy as np
import pandas as pd
import scipy.stats as stats
import math
sales_data = pd.read_excel("extraFilter.xlsx")
valid_sales = sales_data[sales_data["SaleValidity"] <= 2]

def single_var_mean():
    col_name = input('What column would you like to analyze? ')
    current_var = valid_sales[col_name]
        
    

    #current_var = current_var.loc[(current_var != 0).any(axis=1)] removing 0s?

    #current_var.dropna(axis=0)

    # mean and sd
    mean = np.mean(current_var)
    sd = np.std(current_var)

    # confidence interval
    dof = len(current_var) - 1
    alpha = 0.95
    stan_err = sd/math.sqrt(len(current_var))
    t_crit = np.abs(stats.t.ppf((alpha)/2,dof))

    lower_ci = mean - (t_crit*stan_err)
    upper_ci = mean + (t_crit*stan_err)

    confidence = (round(lower_ci,2), round(upper_ci,2))

    print('mean:', round(mean, 2), "\nstandard deviation:", round(sd, 3))
    print('95% confidence of mean:', confidence)

def single_var_median(ci, p): # https://github.com/minddrummer/median-confidence-interval/blob/master/Median_CI.py
    col_name = input('What column would you like to analyze? ')
    data = valid_sales[col_name]
    median = np.median(data)
    if col_name == 'FND':
        print('median', (data[17552] + data[17553]) / 2)

    if type(data) is pd.Series or type(data) is pd.DataFrame:
	# 	#transfer data into np.array
        data = data.values

	#flat to one dimension array
    data = data.reshape(-1)
    data = np.sort(data)
    N = data.shape[0]
	
    lowCount, upCount = stats.binom.interval(ci, N, p, loc=0)
	#given this: https://onlinecourses.science.psu.edu/stat414/node/316
	#lowCount and upCount both refers to  W's value, W follows binomial Dis.
	#lowCount need to change to lowCount-1, upCount no need to change in python indexing
    lowCount -= 1
    confidence = (data[int(lowCount)], data[int(upCount)])
	# print lowCount, upCount
    print('95% confidence: ', confidence, 'median:', median)

stat = input('Which statistic would you like to analyze (mean or median)? ')
if stat=='mean':
    single_var_mean()
else:
    single_var_median(0.95, 0.5)
