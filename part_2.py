import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import pandas as pd

skip = [0, 2, 3, 22, 26, 27, 35, 44, 52, 67, 71, 72, 73, 74, 76, 87, 99, 100]
data_zpl = pd.read_excel("./xlsx/tab4_zpl_2023.xlsx", sheet_name="с 2018", skiprows=skip)

plt.figure(figsize=(8, 6))
ax = sns.boxplot(y=data_zpl[2020], data=data_zpl, color="red", saturation=0.65, medianprops={"linewidth": 2})
ax.axes.set_title("Среднемесячная номинальная начисленная заработная плата работников по полному кругу организаций в целом по экономике\n по субъектам Российской Федерации 2020 года, рублей", fontsize=14)
ax.set_xlabel("Зарплата по субъектам РФ, 2020 год", fontsize=12)
ax.set_ylabel("Зарплата, руб.", fontsize=12)

def find_outliers(data):
	q1, median, q3 = np.percentile(data, [25, 50, 75])
	ax.annotate(f"Q1 = {q1}", xy=(0, q1), xytext=(3, 3), textcoords="offset points")
	ax.annotate(f"Q3 = {q3}", xy=(0, q3), xytext=(3, 3), textcoords="offset points")
	iqr = q3 - q1
	ax.annotate(f"median = {median}, iqr = {iqr}", xy=(0, median), xytext=(3, 3), textcoords="offset points")
	lower_bound = q1 - (1.5 * iqr)
	upper_bound = q3 + (1.5 * iqr)
	outliers = data[(data < lower_bound) | (data > upper_bound)]
	return outliers

outliers = []
subject_data = data_zpl[2020]
outliers.extend(find_outliers(subject_data))

for outlier in outliers:
	subject = data_zpl[data_zpl[2020] == outlier]["Unnamed: 0"].iloc[0]
	ax.annotate(f"{subject}", xy=(0, outlier), xytext=(3, 3), textcoords="offset points")

range = np.max(subject_data) - np.min(subject_data)
variance = np.var(subject_data)
std_dev = np.std(subject_data)

ax.annotate(f'Range: {range:.2f}', xy=(0.01, 0.9), xycoords='axes fraction', color='red', fontstyle='italic')
ax.annotate(f'Variance: {variance:.2f}', xy=(0.01, 0.85), xycoords='axes fraction', color='green', fontstyle='italic')
ax.annotate(f'Standard Deviation: {std_dev:.2f}', xy=(0.01, 0.8), xycoords='axes fraction', color='blue', fontstyle='italic')

plt.show()