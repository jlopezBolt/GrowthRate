import untangle, argparse
import sys
from argparse import ArgumentParser
from collections import defaultdict
import numpy
from openpyxl import Workbook
from openpyxl.charts import LineChart, Reference, Series, ErrorBar, ScatterChart
import xlsxwriter
import re
import os.path
import collections
import xml.etree.ElementTree as ET
import matplotlib.pyplot as plt
import re
import pylab
from mpl_toolkits.mplot3d import Axes3D
import scipy.stats as t
import numpy as np
from sklearn import linear_model
import math
from math import log
import csv
import itertools
from gooey import Gooey, GooeyParser
from itertools import groupby
import datetime
from itertools import takewhile
import xlrd
from xlutils.copy import copy
from xlsx2csv import *

################################################################################# 
def time_select(x_values, logy_values):
	time_start = 0
	time_end = 3
	time_max = len(x_values)
	r_value_list = []
	slope_list = []
	intercept_list = []
	p_value_list = []
	std_err_list = []
	time_start_list = []
	time_end_list = []
	while time_end <= time_max:
		slope, intercept, r_value, p_value, std_err = t.linregress(x_values[time_start:time_end], logy_values[time_start:time_end])
		slope_list.append(slope)
		intercept_list.append(intercept)
		r_value_list.append(r_value)
		p_value_list.append(p_value)
		std_err_list.append(std_err)
		time_start_list.append(time_start)
		time_end_list.append(time_end)
		time_start = time_start+1
		time_end = time_end+1
	
	parameter_array = zip(slope_list, intercept_list, r_value_list, p_value_list, std_err_list, time_start_list, time_end_list)
	
	for x in parameter_array:
		if max(slope_list) == x[0]:
			parameter_out = x
	return parameter_out

################################################################################# 

 
def average(values): 
	average = []
	average = sum(values)/len(values)
	return average
	
################################################################################# 
	
def plate_averages(well_value, plate_coordinate,input):
	value_ave = [well_value[x:x+96] for x in xrange(0, len(well_value), 96)]
	well_ave = [plate_coordinate[x:x+96] for x in xrange(0, len(plate_coordinate), 96)]

	num_plates = len(value_ave)
	num_strains = len(input[2])
	strain_counter = 1
	time_points = input[0]
	averages_out = []
	stdv_out = []
	while strain_counter < num_strains:
		averages = []
		blank_wells = input[3]
	
		plate_counter = 0
		current_strain = input[2][strain_counter]
		current_strain_wells = input[strain_counter+3]
		strain_averages_out = []
		strain_stdv_out = []
		corrected_strain_values =[]
		while plate_counter < num_plates:
			current_time = time_points[plate_counter]
			current_plate_dil = float(input[1][plate_counter])
		
			current_plate_data = map(None, well_ave[plate_counter], value_ave[plate_counter])
		
			plate_blank_values = []
			for x in current_plate_data:
				if x[0] in blank_wells:
					plate_blank_values.append(float(x[1]))
			average_blank_value = average(plate_blank_values)

			counter = 0
			plate_strain_values = []
			for x in current_plate_data:
				if x[0] in current_strain_wells:
					plate_strain_values.append(float(x[1]))
					
			corrected_strain_values_temp = []
			for x in plate_strain_values:
				corrected_strain_values_temp.append(float((x-average_blank_value)*current_plate_dil))

			corrected_strain_values.append(corrected_strain_values_temp)
	
			current_strain_average = average(corrected_strain_values_temp)
			current_strain_stdv = np.std(corrected_strain_values_temp)
			
			current_strain_out = (current_strain, current_time, current_strain_average, current_strain_stdv)
			
			strain_averages_out.append(current_strain_out)
			
			plate_counter = plate_counter+1
			
		averages_out.append(strain_averages_out)
		
		strain_counter = strain_counter+1
	return averages_out


 
################################################################################# 

@Gooey(program_name = "growthrates3")
def parse_args():
	parser = GooeyParser(description="Growth Curve Analysis")
	parser.add_argument('-f', '--elisa_file', help="xml file from plate reader", type=str, default=None, required=True, metavar="File", widget='FileChooser')
	#parser.add_argument('-o', '--out_file', help="name of output xlsx file", type=str, default=None, required=True, widget='DirChooser')
	parser.add_argument('-i', '--data_info', help = "file with data information", type=str, default = None, required = True, widget = 'FileChooser')
	#parser.add_argument('-c', '--graphs', help="name of output xlsx file for graphs", type=str, default=None, required=True)
	#parser.add_argument('-d', '--out_directory', help="directory for outfiles: platemap & correctedOD", type=str, default=None, required = True, widget = 'DirChooser')

	args = parser.parse_args()

	if os.path.isfile(args.elisa_file) == True:
		try:
			data_obj = untangle.parse(args.elisa_file)
		except Exception:
			print("The file %s is not readable!" % args.elisa_file)
			sys.exit(0)
		plate_name = data_obj.Experiment.PlateSections[1].PlateSection['Name']
		return args

args = parse_args()

################################################################################# 


reader_input_template = csv.reader(open(args.data_info, 'rU'))

input = []
for row in reader_input_template:
    input.append(row)

input_array = []
for sublist in input:
    sublist = filter(None, sublist)
    input_array.append(sublist)
 
################################################################################# 
tree = ET.parse(args.elisa_file)
root = tree.getroot()
#out_dir=args.out_directory
#out_file = args.out_file
plate_names = {}

plate_names = defaultdict(list)

################################################################################# 


reader_input_template = csv.reader(open(args.data_info, 'rU'))
for row in reader_input_template:
    input.append(row)

for sublist in input:
    sublist = filter(None, sublist)
    input_array.append(sublist)

time_points =[]   
for sublist in input_array[0]:
    time_points.append(sublist)

all_wells = []
counter_all = 3
while counter_all < len(input_array):														## creates a nested list with each sublist containing the name of the wells for each strain and controls
	all_wells.append(input_array[counter_all])
	counter_all = counter_all+1

strain_name_list = []
for sublist in input_array[2]:       
    strain_name_list.append(sublist)

strain_list = []
for sublist in input_array[2]:
	strain_list.append(sublist)

blank_wells = []
for sublist in input_array[3]:
    blank_wells.append(sublist)

counter = 4
strain_wells = []
while (counter < (len(input_array))):
    strain_wells.append((input_array[counter]))
    counter = counter+1

dilution = []
for sublist in input_array[1]:
	dilution.append(sublist)

dilution_factor =[]
temp1 = []
dilution = list(map(float, dilution))
for sublist in dilution:
	temp1 = [sublist] * 96
	dilution_factor.extend(temp1)


hour = []
temp =[]
time_points = list(map(float, time_points))
for sublist in time_points:
    temp = [sublist] * len(input_array[3])
    hour.extend(temp)

######################################################################

#extracts data values from plates
well_value =[]
well_name = []
plate_name = []
read_time = []
raw_times = []
	
for plate in root.getiterator('PlateSection'):
	raw_times.append(plate.attrib.get('ReadTime'))
	for info in plate.getiterator('Wavelengths'):
		for well in info:
			for node in well:
				for data in node:
					for value in data:	
						read_time.append(plate.attrib.get('ReadTime'))
						plate_name.append(plate.attrib.get('Name'))
						well_value.append(value.text)
						well_name.append(data.get('Name'))

 

well_time = []
for x in read_time:																			## creates a list that has the read time for each well
	well_time_temp = datetime.datetime.strptime(x,'%I:%M %p %m/%d/%Y')
	well_time.append(well_time_temp)

												
matrix = [plate_name, well_name, well_value, dilution_factor]
data = zip(*matrix)

 
################################################################################# 

well_value = [float(i) for i in well_value]
plate_averages_out = plate_averages(well_value, well_name, input_array)

print ''
for i in plate_averages_out:   # each i is a list of plate averages for an individual strain
	strain_y_values = []
	strain_x_values = []
	strain_name_value = []
	strain_logy_values = []
	for y in i:
		strain_y_values.append(float(y[2]))
		strain_logy_values.append(np.log(y[2]))
	for x in i:
		strain_x_values.append(float(x[1]))
	for name in i:
		strain_name_value.append(name[0])
	
	print strain_x_values[2:5]
	print strain_y_values[2:5]
	print strain_logy_values
	
	linregress_out = time_select(strain_x_values, strain_logy_values)
	print linregress_out
	print ''










 ################################################################################# 
"""
for wells in strain_wells:																	# wells is a list of well coordinates for given strain
	od_value = []
	A600_value = []
	log_od_value = []
	for sublist in data:															# sublist in data is list [time stamp, well, value, dilution factor
		if sublist[1] in wells:
			A600_value.append(float(sublist[2]))
			#print sublist[1]															# iterates over individual wells found in well coordinate list of strain
			od_value.append((float(sublist[2])-blank_average)*(sublist[3]))					# calculates the corrected value and adds it to od_value
			log_od_value = [np.log(y) for y in od_value]									# Takes the log of the corrected OD value 
	A600_value_out.append(A600_value)
	od_value_out.append(od_value)
	

	
	log_od_value_out.append(log_od_value)
	max_info, hour_max, time_start  = time_select(hour, log_od_value, input_array)

	time_constraint = max_info[2]
	
	time_used.append(max_info[2])
	
	time_1 = max_info[5]
	hour_max_list.append(hour_max)
	
	hour_set = []
	for element in hour:
		if element not in hour_set:
			hour_set.append(element)
 		
	slope, intercept, r_value, p_value, std_err = t.linregress(hour[time_1:time_constraint], log_od_value[time_1:time_constraint])				# Linear regression using timepoints in hour and the log OD values. Returns regression estimates and inputeters
	

	y_error = y_err1(hour[time_start:time_constraint], slope, intercept,log_od_value[time_start:time_constraint])											# Calculates error in y
	p_x, confs = conf_calc(hour[time_start:time_constraint], y_error, 0.975)																		# Calculates Confidence values for different values of x
	p_y, lower, upper = ylines_calc(p_x, confs, slope, intercept)	
	plot_linreg_CIs(hour[time_start:time_constraint], log_od_value[time_start:time_constraint], p_x, p_y, lower, upper, strain_name_list[counter+1])
	slope_list.append(slope)
	r_value_list.append(r_value)
	std_err_list.append(std_err)
	counter = counter + 1



"""

