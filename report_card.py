import os
import pandas as pd
import numpy as np
import re

path = os.getcwd()
file_path = path+"/input_folder/"
files = os.listdir(file_path)

files.sort()
# sorted(files)

files_csv = [f for f in files if f[-4:] == ".csv"]

# # Read CSV files from List
# df = pd.concat(map(pd.read_csv, files_csv))

# print(files_csv)
# quit()

appended_data = []
for file in files_csv:
	book = pd.read_csv(file_path+file, index_col=False, delimiter=',', error_bad_lines=False)
	df = pd.DataFrame(book)
	df.memory_usage(index=False, deep=True)
	df = df.dropna(how='all')
	df = df.replace(np.nan, "")

	maximum_marks = df.iloc[:, 2].replace("Absent", 0).astype(int)
	highest_marks = maximum_marks.max()

	marks_optain_column = df.columns[2]
	res = marks_optain_column[marks_optain_column.find('(')+1:marks_optain_column.find(')')].split(":")
	out_of_marks = res[1].strip()

	df['date'] = '07.09.22'
	df['type'] = 'DCT'
	df['topic'] = 'Trigonometry'
	df['marks obtained'] = df.iloc[:, 2]
	df['out of'] = out_of_marks
	df['highest marks'] = highest_marks

	cols = df.columns.tolist()
	cols = ['date', 'Name', 'type', 'topic', 'marks obtained', 'out of', 'highest marks']
	df = df[cols]

	# final_df = pd.concat(book, ignore_index=True)
	appended_data.append(df)

	# print(df)
	# quit()
	
appended_data = pd.concat(appended_data)
appended_data = appended_data.sort_values('Name')

for i, x in appended_data.groupby('Name'):
	p = os.path.join(os.getcwd(), "data_{}.csv".format(i.lower()))
	x.to_csv(p, index=False)
	
# print(appended_data)
quit()


