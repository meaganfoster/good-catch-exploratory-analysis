# Import packages
import nltk
from nltk.stem.snowball import SnowballStemmer
from nltk.corpus import stopwords
import glob
import os
import pandas as pd
import numpy as np
import scipy
import csv
import re
import sklearn
from sklearn.linear_model import Ridge
from sklearn.metrics import mean_squared_error
from sklearn.feature_extraction.text import TfidfVectorizer, CountVectorizer
import spacy
import en_core_web_sm
nlp = en_core_web_sm.load()
nlp = spacy.load("en_core_web_sm")           # load package "en_core_web_sm"
from collections import Counter


## 1 ## Import files

## Import 1 excel file
# filepath = 'gc_Jan12_Dec12_main.xlsx'
# dataframe = pd.read_excel(filepath)

# Import multiple excel files from the "Data for Import" folder
filepath = r'C:\Users\mlfos\OneDrive - University of North Carolina at Chapel Hill\Coursework\INLS690\Good catches\Data for Import' # use your path
files = os.listdir(filepath)
print("The following files are in the folder: ")
print(files)
print("Only files with .xlsx file extensions will be imported.")


data_frame = pd.DataFrame()
# for files in glob.glob("*.xlsx"):
#    data = pd.read_excel(os.path.join(filepath, file))
for file in files:
   if file.endswith('.xlsx'):
    print("reading " + file + " ...")
    data = pd.read_excel(os.path.join(filepath, file))
    data_frame = data_frame.append(data, ignore_index=True)
    print(file + " has been appended to the dataframe.")
# print("The final dataframe: ")
# print(data_frame)

#Drop unwanted columns
data_frame_trimmed =  data_frame.iloc[: , 1:]
data_frame_trimmed = data_frame_trimmed.drop(columns=['GC_Label', 'FQID', 'QSC', 'Status', 'Champ', 'Publish Option', 'Closed', 'Closed At', 'Closed By', 'Ignored', 'Resolved', 'Auto Updated',
                                              'Feedback Status', 'Marked For Review', 'No Care Path', 'Care Path', 'Tripped Step', 'Caught Step', 'T: Not on CP/Other QSC/Approx./Unknown',
                                              'C: Not on CP/Other QSC/Approx./Unknown', 'Tripped from QSC', 'Caught by QSC'])
print(data_frame_trimmed)
# Include all available data, i.e. no training/validation split
train_dataframe = data_frame_trimmed.copy()
print('Final dataframe contains', len(train_dataframe), 'rows.')

## 2 ## Pre-processing
print()
print("Begin pre-processing...")
# Remove special characters
train_dataframe['Description_SpecChar'] = train_dataframe['Description'].str.replace('\W', ' ', regex=True)
print("Remove special characters, complete.")

# Convert all text to lowercase
train_dataframe['Description_lower'] = train_dataframe['Description_SpecChar'].str.lower()
print("Convert to lowercase, complete.")

# Remove the stopwords
stop_words = set(stopwords.words('english'))
train_dataframe['Description_noSW'] = train_dataframe['Description_lower'].apply(lambda x: ' '.join(word for word in x.split() if word not in stop_words))
print("Stopword removal complete.")

# Use English stemmer.
stemmer = SnowballStemmer("english")
train_dataframe['Description_stemmer'] = train_dataframe['Description_noSW'].apply(lambda x: " ".join(stemmer.stem(word) for word in x.split()))
print("Stemming complete.")

# Print and save final data frame
print("Pre-processing complete. File saved as 'train_dataframe_clean.csv'")
train_dataframe.to_csv('train_dataframe_clean.csv', sep='\t')



## 4 ## Analyze words
print("Begin tokenization...")
# Create token of each word in the description of each row
tokens = train_dataframe["Description_stemmer"].str.strip().str.split()
# print(tokens)

# convert to list of lists
word_list = [w for w in tokens]
# print(word_list)

# flatten word list
word_list_flat = [x for l in word_list for x in l]
# print(word_list_flat)

# get distinct count of tokens/words
c = Counter(word_list_flat)
print("Unigram word frequency list:")
print(c)

# Save distinct word count and freq to csv file
c_2 = pd.DataFrame.from_dict(c, orient='index').reset_index()
c_2.to_csv('unigram_word_list_frequency.csv', sep='\t')
print("Unigram frequency complete. See 'unigram_word_list_frequency.csv'")

# find consecutive bi-grams
bigram_word_list = (pd.Series(nltk.ngrams(word_list_flat, 2)).value_counts()) # [:10]
print(bigram_word_list)




# Save distinct bigram count and freq to csv file
c_3 = bigram_word_list.to_csv('bigram_word_list_frequency.csv', sep='\t')
print(c_3)

# Apply POS tagger
from nltk import word_tokenize, pos_tag, pos_tag_sents

train_dataframe['POSTags'] = pos_tag_sents(train_dataframe['Description_stemmer'].apply(word_tokenize).tolist())
print(train_dataframe)
POS_tag_df = train_dataframe['POSTags']
POS_tag_df = POS_tag_df.explode('POSTags')
print(POS_tag_df)

from collections import Counter
c = Counter(POS_tag_df)
print(c)

POS_tag_df2 = pd.DataFrame(POS_tag_df.tolist(), columns=["orig_word", "POS_tag"], index=POS_tag_df.index)
POS_tag_df2.to_csv('POSTags_tocolumns.csv', sep='\t')


# orig_words_freq_counts = POS_tag_df2.groupby(["orig_words"]).counts()
# $print(orig_words_freq_counts)

POS_tag_df_final = POS_tag_df2.drop_duplicates(subset=["orig_word", "POS_tag"], keep='first')
print(POS_tag_df_final)
POS_tag_df_final.to_csv('POSTags_deduped.csv', sep='\t')




## ## Identify clusters using semi-processed text

from sklearn.feature_extraction.text import TfidfVectorizer
vectorizer = TfidfVectorizer()
X = vectorizer.fit_transform(train_dataframe['Description_lower'])
print("Vectorizer applied.")

import matplotlib.pyplot as plt
from sklearn.cluster import KMeans
Sum_of_squared_distances = []
K = range(2, 10)
for k in K:
   km = KMeans(n_clusters=k, max_iter=200, n_init=10)
   km = km.fit(X)
   Sum_of_squared_distances.append(km.inertia_)
plt.plot(K, Sum_of_squared_distances, 'bx-')
plt.xlabel('k')
plt.ylabel('Sum_of_squared_distances')
plt.title('Elbow Method For Optimal k')
plt.show()

true_k = 6
model = KMeans(n_clusters=true_k, init='k-means++', max_iter=200, n_init=10)
model.fit(X)
labels=model.labels_
report_clusters=pd.DataFrame(list(zip(train_dataframe['Description_lower'], labels)), columns=['title', 'cluster'])
print(report_clusters.sort_values(by=['cluster']))
print("Clustering complete.")

from wordcloud import WordCloud
result = {'cluster': labels, 'wiki': train_dataframe['Description_lower']}
result = pd.DataFrame(result)
exit()
for k in range(0,true_k):
   s=result[result.cluster==k]
   text=s['wiki'].str.cat(sep=' ')
   text=text.lower()
   text=' '.join([word for word in text.split()])
   wordcloud = WordCloud(max_font_size=50, max_words=100, background_color="white").generate(text)
   print('Cluster: {}'.format(k))
   print('Titles')
   titles=report_clusters[report_clusters.cluster==k]['title']
   print(titles.to_string(index=False))
   plt.figure()
   plt.imshow(wordcloud, interpolation="bilinear")
   plt.axis("off")
   plt.show()


exit()

# evaluate an ridge regression model on the dataset
from numpy import mean
from numpy import std
from numpy import absolute
from pandas import read_csv
from sklearn.model_selection import cross_val_score
from sklearn.model_selection import RepeatedKFold
from sklearn.linear_model import Ridge
# load the dataset
# data = data.values
# X = vectorizer.transform(data[0:1]).toarray()
# y = vectorizer.transform(data[:, -1])
# X, y = data[:, :-1], data[:, -1]
# print(X, y)
# define model
model = Ridge(alpha=1.0)
# model = Ridge(alpha=1).fit(X, y)
# define model evaluation method
cv = RepeatedKFold(n_splits=10, n_repeats=3, random_state=1)
# evaluate model
# scores = cross_val_score(model, X, y, scoring='neg_mean_absolute_error', cv=cv, n_jobs=-1)
# force scores to be positive
# scores = absolute(scores)
# print('Mean MAE: %.3f (%.3f)' % (mean(scores), std(scores)))