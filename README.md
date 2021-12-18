# EE_551
Merge, Sort, analyze NTSB Data

Python 3.10 code for merging, sort, and highlevel analysis of NTSB data
assumes that the data within the excel sheet has been converted to numbers

This script requires the following two files:
stopwords_en.txt <- extra stop words to remove 
2010_NTSB_avail_narratives_python.xlsx  <- filtered for 2010, and removed chars :, |, \, /


Code also requires the following modules 
import xlrd #to read excel workbooks
import pandas as pd #used to write new excel sheet from user input and box plots
import openpyxl as op #for exporting excel columns into a text files
import nltk #for text part of code
import string  #for text part of code
import numpy as np  #for text part of code and box plots
import matplotlib.pyplot as plt  #for plotting part of code
import os    #for ploting bar charts
from os import path  #for ploting bar charts
from wordcloud import WordCloud  #for word clouds
from pandas import ExcelWriter #reading/writing excel files
from pandas import ExcelFile   #reading/writing excel files

