#Python 3.10 code for merging, sort, and highlevel analysis of NTSB data
#assumes that the data within the excel sheet has been converted to numbers

#This script requires the following two files:
#stopwords_en.txt <- extra stop words to remove 
#2010_NTSB_avail_narratives_python.xlsx  <- filtered for 2010, and removed chars :, |, \, /


#First importing all the modules required for this code
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

print ('Ok, all the modules have been imported.')

#==========Code below is to gather 2 excel sheets and merge them =================================
print ('Next we will merge two NSTB excel sheets.')
FirstSheet = input('Enter the first excel sheet that you want merged, and include the .xlsx extention: ')
Sheet1 = pd.read_excel(FirstSheet)
SecondSheet = input('Enter the second excel sheet that you want merged, and include the .xlsx extention again: ')
Sheet2 = pd.read_excel(SecondSheet)
print('Thanks! Time for me to merge them....')

# merging the two sheets by the first two columns
result = pd.merge(Sheet1, Sheet2, how="outer", on=["ev_id", "Aircraft_Key"], left_index=False)
result.to_excel('Merged_Sheets.xlsx')
merged = "Merged_Sheets.xlsx"
print ('Good news, the merged sheet has been created and is called Merged_Sheets.xlsx')


#==========Code below to sort on the merged excel sheets=================================

#Code below is to read and get the excel info
inputWorkbook = xlrd.open_workbook(merged)
inputWorksheet = inputWorkbook.sheet_by_index(0)

#Columns needs to be decremented so that the index column is not counted 
print('Your merged excel sheet has', inputWorksheet.ncols-1, 'number of columns.')
print('And has', inputWorksheet.nrows, 'number of rows.')

#Creating a variable for the column, so that the column count is accourate
#The count starts at zero with python and we want to omit the first two
column = inputWorksheet.ncols-2

#list of the attributes in first row of the workbook
attributes = []

#Create a readable list for the user on what are the columns in the excel sheet
for y in range(1, inputWorksheet.ncols):
    attributes.append(inputWorksheet.cell_value(0, y))
    
print("These are the columns available in this sheet:")

#list out each attributes with the column number as a list so it is easier to read
var_holder = {}
for i in range(column):
    var_holder['column' + str(i)] = "Attribute"+attributes[1]
    print("column", i+1, "is", attributes[i])

#Ask for the item to sort on 
SelectedAttribute = input("Enter the attribute that you want to filter on: ")
print("you selected:", SelectedAttribute)
print("now getting all the values for", SelectedAttribute)

#to get unquie items in a column selected
ItemsInSelectedAttribute = pd.read_excel(merged)
print("these are the values for:", SelectedAttribute)
print((ItemsInSelectedAttribute[SelectedAttribute].unique()))

#Ask for the value within that item to sort on  
SelectedValue = input("Enter the attribute value that you want to filter on: ")
print("you selected ", SelectedValue)

#selectedValue must be a number to work for filtering!
#changing SelectedValue to an integer
SV=int(SelectedValue)

print("Now we need to name your new workbook that will contain only items that are filtered on ", SelectedValue, "from column ", SelectedAttribute)
NewWorkbook = input("Enter the new workbook name:")
NewExcelFile = NewWorkbook+'.xlsx'
print("you named this file:", NewExcelFile)

#sorting and creating new file for user
df = pd.read_excel(merged, index_col=None)
df = df.loc[df[SelectedAttribute] == SV] #this is the selectedValue

writer = pd.ExcelWriter(NewExcelFile)
df.to_excel(writer, index=False) 
writer.save()

print('file', NewExcelFile, 'has been created')

#==========Code below gather narratives from the sorted excel sheet=================================

print("Now I will get the related narratives available for your selection.")
#getting the narrative file that contains the narratives
narpath = '2010_NTSB_avail_narratives_python.xlsx'
#and it will be compaired with the one created earilier
path = NewExcelFile

SortedBook = pd.read_excel(path)
#need to list the max possible rows of the sheet
df1 = SortedBook.head(900000)
  
#Converting first column to list using list()
SortedBook_EV_ID_List = list(df1['ev_id']) 

#Pulling the data for the narative excel sheet
NarBook = pd.read_excel(narpath)
#for max rows of the sheet possible Number of rows to select.
df2 = NarBook.head(100000)
  
#Converting the narative first column to the same list 
NarBook_EV_ID_List = list(df2['ev_id']) 
print("I have all the data, now I am compairing...")

#Compairing both lists and will only keep the ones matching with the &
EV_ID_matching = list(set(NarBook_EV_ID_List) & set(SortedBook_EV_ID_List))

#Make a newNarrativeSheet, first need to call out the narr copy first column 'ev_id'
print("Now we need to name your narrative workbook that will contain only your items.")
NewNarWorkbook = input("Enter the new workbook name (I'll add the .xlsx for you): ")
NewNarExcelFile = NewNarWorkbook+'.xlsx'
print("Great! I have ", NewNarExcelFile)

#writting the data and saving the new file
df2 = df2[df2['ev_id'].isin(EV_ID_matching)]
writer = pd.ExcelWriter(NewNarExcelFile)
df2.to_excel(writer, index=False) 
writer.save()

print("The new file", NewNarExcelFile, "has been created.")

#==========Code below write narratives to a text file=================================

print("Now I will create a text file of just the narratives.")

#First need to read the new excel just created with the naratives in Column C
NarSortedBook = pd.read_excel(NewNarExcelFile, index_col=0)
#for max rows of the sheet possible Number of rows to select.
df1 = NarSortedBook.head(900000)
df1 = pd.DataFrame(data=NarSortedBook)

inputNarWorkbook = xlrd.open_workbook(NewNarExcelFile) 
inputNarWorksheet = inputNarWorkbook.sheet_by_index(0)
#bookrows used for counting later
BookRows = inputNarWorksheet.nrows-1

print("Now creating a text file for your data...")

#UTF-8 needed incase there are some odd charactors in the NTSB data
Factual_Narrative = open('Factual_Narrative.txt','w', encoding='UTF-8')

#for loop to read through the sheet
for y in range(BookRows): 
    cellData1 = df1.iloc[y, 1] #df1.iloc[row, column] so this is all rows for colmun C
    Factual_Narrative.write(cellData1+'\n')
    
Factual_Narrative.close()

print("Factual_Narrative.txt has been created")
print("Time to clean this file!")


#===============Clean the text file==================


#using UTF8 again because there may be unknown charators that will read 
file = open('Factual_Narrative.txt', 'rt', encoding='UTF8')
text = file.read()
file.close()
print("As I clean this file, I will show you the first ten words with each step so you can see me in action...")

#split into words by white space
print('split into words by white space')
words = text.split()
print(words[:10])

# split into sentences with nltk
print('split into sentences with nltk')
from nltk import sent_tokenize
sentences = sent_tokenize(text)
print(sentences[:10])

# split into words with nltk
print('split into words with nltk')
from nltk.tokenize import word_tokenize
tokens = word_tokenize(text)
print(tokens[:10])

# remove all tokens that are not alphabetic with nltk
print('remove all tokens that are not alphabetic with nltk')
words = [word for word in tokens if word.isalpha()]
print(words[:10])

#Filter out Stop Words
print('Filter out Stop Words')
from nltk.corpus import stopwords
stop_words = stopwords.words('english')
#print(stop_words)

# convert to lower case
print("Converting everything to lower case (NO YELLING!)")
tokens = [w.lower() for w in tokens]
print(words[:10])

# remove punctuation from each word
print("Getting rid of punctuation from each word")
import string
table = str.maketrans('', '', string.punctuation)
stripped = [w.translate(table) for w in tokens]
print(words[:10])

# remove remaining tokens that are not alphabetic
print("Remove remaining items that are not alphabetic")
words = [word for word in stripped if word.isalpha()]
print(words[:10])

# filter out stop words
print("Removing common words")
from nltk.corpus import stopwords
stop_words = set(stopwords.words('english'))
words = [w for w in words if not w in stop_words]
print(words[:10])

# filter my own stop words that are not useful for NTSB data, like airplane
myfile = open('stopwords_en.txt', 'rt')
mytext = myfile.read()
myfile.close()
print("Time to remove my own stop words!")
words = [w for w in words if not w in mytext]
print(words[:10])

#last version of the file is called words
print("Now we are all clean and making this a new file for you.")
CleanTxtFileName = input("What do you want the name of this file to be called (I'll add the .txt for you)?")
CleanTxtFile = CleanTxtFileName+'.txt'
print("Your new file being created is:", CleanTxtFile)

with open(CleanTxtFile, "w") as output:
    output.write(str(words))
print(CleanTxtFile, "is available to view.")


#===============Show/get data of the words from the clean the text file==================

#Open text file and start gathering data from it to show
CleanedTextFile = open(CleanTxtFile, 'r')
read_data = CleanedTextFile.read()
per_word = read_data.split()

print("The total words in this file are: ", len(per_word))

count = 0
Unquiewords = set(read_data.split())
for word in Unquiewords:
    count += 1
    
print("And the total unique words are: ", count)

#Now getting the most occuring words for plotting
from collections import Counter
def count_word(CleanTxtFile):
        with open(CleanTxtFile) as f:
                return Counter(f.read().split())

#Showing top 20 words
TopOccurances = count_word(CleanTxtFile)
Top20 = TopOccurances.most_common(20)
print("The top 20 most frequent words are :", Top20)

#asking the user to search for a word they are interested in
occurrences = input("Enter a word that you want me to count: ")
word_count = words.count(occurrences)
print(f"The word ", occurrences, " appeared ", word_count, " times.")


#Need to flip the array for the plot
Top20Array = np.array(Top20)
result = np.flip(Top20Array)
 
#getting the data for the top 20 for the chart
TopWords = result[ :,0]
Quanity = result[ :,1]

# Creating a simple bar chart for the top 20 words
plt.figure(figsize=[18,6])
col_map = plt.get_cmap('tab20')
plt.bar(Quanity, TopWords, width=0.5, color=col_map.colors, edgecolor='k', linewidth=2)

plt.title('The Top 20 Occuring Words')
plt.xlabel('Words', fontsize=6)
plt.ylabel('Occurance', fontsize=6)
plt.show()

#===============Word Cloud for the text file==============

print("Lets see all your words now")
#to get the data directory
d = os.getcwd()

#Generate a word cloud image
wordcloud = WordCloud().generate(read_data)

# Display the generated image
import matplotlib.pyplot as plt
plt.imshow(wordcloud, interpolation='bilinear')
plt.axis("off")

# max font size at 30
wordcloud = WordCloud(max_font_size=30).generate(read_data)
plt.figure()
plt.imshow(wordcloud, interpolation="bilinear")
plt.axis("off")
plt.show()

print("All done! Have a nice day!")


