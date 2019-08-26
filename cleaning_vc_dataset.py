#!/usr/bin/env python
# coding: utf-8

# In[1]:

# import modules
import pandas as pd
import numpy as np

# this function is used later on to format data. e.g. 75k converted to 75000 etc.
def convert_to_number(x):
    val = 0
    if x is np.nan:
        return np.nan
    elif 'k' in x or 'K' in x:
        val = round(float(x.replace('k', '').replace('K','')) * 1000, 2) # convert k or K  to a thousand
    elif 'm' in x or 'M' in x:            
        val = round(float(x.replace('m', '').replace('M','')) * 1000000, 2) # convert m or M to a million
    elif 'b' in x or 'B' in x:
        val = round(float(x.replace('b', '').replace('B','')) * 1000000000, 2) # convert b or B to a Billion
    else:
        val = round(float(x), 2) 
    print(val)
    return val

# In[2]:
# read csv file data into a dataframe
df = pd.read_csv('vc1.csv') 

# In[3]:
# initial preview of first five rows
df.head()

# In[4]:
# remove last unnamed column which is not part of our dataset
df.drop(df.columns[len(df.columns)-1], axis=1, inplace=True) 

# In[5]:
# replace spaces and empty fields with NaN
df.replace(r'\s+', np.nan, regex=True).replace('',np.nan)

# dataframe after dropping last column
df.head()

# In[6]:
# number of row entries and columns in our dataframe
df.shape

# In[7]:
# convert all strings to lower case
df = df.applymap(lambda s:s.lower() if type(s) == str else s)

# In[8]:
# see if strings are converted to lowercase
df.head()

# In[9]:
# missing values in each column - How sparse is our data ?
df.isnull().sum() 

# In[10]:
# data type of each column
df.dtypes

# In[11]:
# the entry at index 778 is a duplicate of above entry at index 777
# also this entry contains garbage values in "Investment Amount Minimum" and some other entries.
# delete this entry in next step
print(df.loc[[778]])

# In[12]:
# drop the above entry
df = df.drop([778,], axis=0)
# reindex the entries
df.index = range(len(df))
df.shape

# In[13]:
# new entry at 778 index
print(df.loc[[778]])

# In[14]:
# Clean data for Minimum, Maximum and Previously Invested Total Amount
# Remove Dollar Symbol 
# df['Investment Amount Minimum'] = df['Investment Amount Minimum'].str.replace('$','').str.replace(',','')
# Above steps not needed as it can be done using find and replace in excel for simplicity

# Replace Million or Billion figures with actual floating numbers
df['Investment Amount Minimum'] = df['Investment Amount Minimum'].apply(lambda x: convert_to_number(x))
df['Investment Amount Maximum'] = df['Investment Amount Maximum'].apply(lambda x: convert_to_number(x))
df['Previously Invested Total Amount'] = df['Previously Invested Total Amount'].apply(lambda x: convert_to_number(x))

# In[15]:
print(df['Investment Amount Minimum'].dtype)
print(df['Investment Amount Maximum'].dtype)
print(df['Previously Invested Total Amount'].dtype)
df.tail()

# In[16]:
# remove exact duplicates
df.drop_duplicates()
# shape of dataframe remains same as there are no exact duplicates
df.shape

# In[17]:
# All duplicate entries by "Name of VC"
df[df.duplicated(['Name of VC'], keep='first')]

# In[18]:
# remove duplicates by "Name of VC" and keep those which have highest "Previously Invested Total Amount"
df.sort_values('Previously Invested Total Amount', ascending=False).drop_duplicates('Name of VC').sort_index().reset_index(drop=True)
df

# In[19]:
# reindex the entries
df.index = range(len(df))

df.shape

# In[20]:
df.head()

# In[21]:
df1 = df.fillna("NaN")
df1.to_csv('cleaned_vc_data.csv', index=False)

# removed some further duplicates which were somehow left out (around 139) from this file using excel