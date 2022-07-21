import os
import imageio.v2 as imageio
import xlsxwriter
from math import ceil

pic_name = 'dog.jpg'
im 		 = imageio.imread(pic_name)

num_of_rows   = im.shape[0]
num_of_cols   = im.shape[1]
num_of_colors = im.shape[2]

colors = ['#FF0000', '#00FF00', '#0000FF']

name, ext = os.path.splitext(pic_name)
workbook  = xlsxwriter.Workbook(f'{name}.xlsx')
worksheet = workbook.add_worksheet()
	
data   = []

for row in range(0, num_of_rows):

	for color in range(0, num_of_colors):
		line = []
		
		for collumn in range(0, num_of_cols):
			line.append( im[row, collumn, color] )
		
		data.append(line)
	
for i, line in enumerate(data):
	worksheet.write_row(i, 0, line)
	worksheet.conditional_format(i, 0, i, 1000, {'type': '2_color_scale',
												'min_value': 0,
												'max_value': 255,
                                         		'min_color': "#000000",
                                         		'max_color': colors[i%3]})

workbook.close()
