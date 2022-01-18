# Imports
import nltk
from nltk.stem.snowball import SnowballStemmer
from nltk.corpus import stopwords
from nltk import word_tokenize, pos_tag, pos_tag_sents
from collections import Counter
from math import *
from wordcloud import WordCloud, STOPWORDS, ImageColorGenerator
import glob
import os
import pandas as pd
pd.set_option('display.max_columns', None)
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
nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')
nltk.download('stopwords')


# Globals
current_directory = os.getcwd()


# 1 # Import
def import_multi_excel_files():
    filepath = r'C:\Users\mlfos\OneDrive - University of North Carolina at Chapel Hill\Coursework\INLS690\Good catches\Data for Import'  # use your path
    files = os.listdir(filepath)
    print("The following files are in the folder: ")
    print(files)
    print("Only files with .xlsx file extensions will be imported.")

    data_frame = pd.DataFrame()
    for file in files:
        if file.endswith('.xlsx'):
            print("reading " + file + " ...")
            data = pd.read_excel(os.path.join(filepath, file))
            data_frame = data_frame.append(data, ignore_index=True)
            print(file + " has been appended to the dataframe.")

    # to manually import one file, use:
    # filepath = 'gc_Jan12_Dec12_main.xlsx'
    # dataframe = pd.read_excel(filepath)
    return(data_frame)

def trim_transform_file(import_df):
    transformed_df = import_df.drop(
        columns=['GC_Label', 'FQID', 'QSC', 'Status', 'Champ', 'Publish Option', 'Closed', 'Closed At', 'Closed By',
                 'Ignored', 'Resolved', 'Auto Updated',
                 'Feedback Status', 'Marked For Review', 'No Care Path', 'Care Path', 'Tripped Step', 'Caught Step',
                 'T: Not on CP/Other QSC/Approx./Unknown',
                 'C: Not on CP/Other QSC/Approx./Unknown', 'Tripped from QSC', 'Caught by QSC'])

    # Convert 'Created At' to datetime
    transformed_df['Created At'] = pd.to_datetime(transformed_df['Created At'])
    return transformed_df


def print_column_stats(transformed_dataframe, dataframe_name):
    record_count = len(transformed_dataframe)
    print()
    print('The are ', record_count, 'reports in the ', dataframe_name, 'data set. File stats and reports below:')

    # Convert 'Created At' to datetime
    transformed_dataframe['Created At'] = pd.to_datetime(transformed_dataframe['Created At'])

    # print("Count by year submitted:")
    # counts_by_year = transformed_dataframe.groupby([transformed_dataframe['Created At'].dt.year]).nunique()
    # print(counts_by_year)

    print()
    print('Report counts by Severity and year Created At.')
    print(pd.crosstab(transformed_dataframe['Created At'].dt.year, transformed_dataframe['Severity']))
    print()
    print('Report counts by Created By and year Created At.')
    print(pd.crosstab(transformed_dataframe['Created By'], transformed_dataframe['Created At'].dt.year))
    print()
    print('Report counts by Submitted By and year Created At.')
    print(pd.crosstab(transformed_dataframe['Submitted By'], transformed_dataframe['Created At'].dt.year))

    print(transformed_dataframe.describe())
    return record_count


def export_to_excel(df_to_excel):
    name = [x for x in globals() if globals()[x] is df_to_excel][0]
    excel_writer = pd.ExcelWriter(current_directory + "\\Output\\" + name + ".xlsx", engine='xlsxwriter')
    df_to_excel.to_excel(excel_writer, index=False)
    excel_writer.save()
    print(name + ".xlsx is saved to the Output folder.")


# # Import all .xlsx files from the "Data For Import" folder
import_df = import_multi_excel_files()

# # Drop unwanted columns and export merged data frame
df = trim_transform_file(import_df)

# # Print file statistics and export to excel
num_reports = print_column_stats(df, 'merged')

# # Export final dataframe to MS excel
imported_files_merged = df.copy()
export_to_excel(imported_files_merged)
# print(imported_files_merged)


# 2 # Pre-processing
def dict_export_to_excel(cluster_df__to_excel, file_name):
    writer_dict_df = pd.ExcelWriter(current_directory + "\\Output\\" + file_name + ".xlsx", engine='xlsxwriter')
    cluster_df__to_excel.to_excel(writer_dict_df, sheet_name='main')
    writer_dict_df.save()
    print(file_name + ".xlsx is saved to the Output folder.")

def perform_preprocessing():
    print()
    print("Begin pre-processing...")

    # # Remove special characters
    df['Description_SpecChar'] = df['Description'].str.replace('\W', ' ', regex=True)
    print("Remove special characters, complete.")

    # # Convert all text to lowercase
    df['Description_lower'] = df['Description_SpecChar'].str.lower()
    print("Convert to lowercase, complete.")

    # # Remove the stopwords
    stop_words = set(stopwords.words('english'))
    df['Description_noSW'] = df['Description_lower'].apply(lambda x: ' '.join(word for word in x.split() if word not in stop_words))
    print("Stopword removal complete.")

    # # Use English stemmer.
    stemmer = SnowballStemmer("english")
    df['Description_stemmer'] = df['Description_noSW'].apply(lambda x: " ".join(stemmer.stem(word) for word in x.split()))
    print("Stemming complete.")

    # # Apply POS tagger
    df['POSTags'] = pos_tag_sents(df['Description_stemmer'].apply(word_tokenize).tolist())
    POS_tag_df = df['POSTags']
    POS_tag_df = POS_tag_df.explode('POSTags')

    # Convert POS results to list
    POS_tag_df2 = pd.DataFrame(POS_tag_df.tolist(), columns=["orig_word", "POS_tag"], index=POS_tag_df.index)
    # Remove duplicates and save to excel
    # pos_deduped = POS_tag_df2.drop_duplicates(subset=["orig_word", "POS_tag"], keep='first')

    # Get distinct counts from POS tagger results, use:
    POS_tagger_result_list = []
    from collections import OrderedDict
    c = Counter(POS_tag_df)
    sorted_dict = OrderedDict(sorted(c.items(), key=lambda kv: kv[1], reverse=True))
    print()
    print("Results of the POS tagger, below:")
    for term in sorted_dict:
        print(term, sorted_dict[term])

    # # Print and save final data frame
    print("Pre-processing complete. Saving data to excel.")
    preprocessed_desc_stepbystep_df = df.copy()
    # export_to_excel(preprocessed_desc_stepbystep_df)

    # Save pre-processed description to its own data frame
    preprocessed_desc = df['Description_stemmer'].copy()
    # to export:
        # df['Description_stemmer'].to_csv(r'C:\Users\mlfos\OneDrive - University of North Carolina at Chapel Hill\Coursework\INLS690\Good catches\Output\preprocessed_desc.txt', sep=' ', header=False, index=False)
        # export_to_excel(preprocessed_desc_df)

    return preprocessed_desc

# # Perform pre-processing steps
preprocessed_desc_df = perform_preprocessing()



# 3 # Analyze words
def process_text_into_word_list(text):
    tokens = text.strip().split()
    word_list = [w for w in tokens if not w.startswith('#') and not w.startswith('@')]
    return word_list


def get_itemset_frequency(report_dataframe, itemset_generator):
    # treat each report as a window of text
    # count how many windows an itemset appears in
    itemset_freq = {}
    for line in report_dataframe:
            itemset_list = itemset_generator(line)
            # accumulate counts
            for itemset in set(itemset_list):
                if itemset not in itemset_freq:
                    itemset_freq[itemset] = 1
                else:
                    itemset_freq[itemset] += 1
    return itemset_freq


def process_text_into_bigram_list(text):
    tokens = text.strip().split()
    bigram_list = []
    for i in range(len(tokens) - 1):
        if not tokens[i].startswith('#') and            not tokens[i].startswith('@') and            not tokens[i+1].startswith('#') and            not tokens[i+1].startswith('@'):
            bigram_list.append( (tokens[i], tokens[i + 1]) )
    return bigram_list


def get_pmi_result(unigram_freq, bigram_freq):
    pmi_result_list = list()
    pmi_per_bigram = {}
    for bigram, freq in bigram_freq.items():
        if bigram[0] in unigram_freq and bigram[1] in unigram_freq:
            # calculate the pointwise mutual information for this bigram
            # you will need to use the following variables:
            #
            # freq: frequency of this bigram
            # unigram_freq[bigram[0]]: frequency of the first word in the bigram
            # unigram_freq[bigram[1]]: frequency of the second word in the bigram
            # num_reports: total number of reports in the data (i.e. transactions in the database)
            #
            pmi_result = {'pmi_per_bigram': 0.0, 'bigram': str, 'freq': 0.0, 'bigram[0]': str, 'unigram_freq[bigram[0]]': 0.0, 'bigram[1]': str, 'unigram_freq[bigram[1]]': 0.0, 'has_word_1_has_word_2': 0.000, 'no_word_1_has_word_2': 0.000, 'has_word_1_no_word_2':0.000,'no_word_1_no_word_2':0.000}
            pmi_result['bigram']=bigram
            pmi_result['freq']=freq
            pmi_result['bigram[0]']=bigram[0]
            pmi_result['unigram_freq[bigram[0]]']=unigram_freq[bigram[0]]
            pmi_result['bigram[1]']=bigram[1]
            pmi_result['unigram_freq[bigram[1]]']=unigram_freq[bigram[1]]


            pmi_result['has_word_1_has_word_2']=freq/num_reports  # a = freq
            pmi_result['no_word_1_has_word_2']=(unigram_freq[bigram[0]]-freq)/num_reports  # b = unigram_freq[bigram[1]]-freq
            pmi_result['has_word_1_no_word_2']=(unigram_freq[bigram[1]]-freq)/num_reports   # c = unigram_freq[bigram[0]]-freq
            pmi_result['no_word_1_no_word_2']=(num_reports-((freq)+(unigram_freq[bigram[1]]-freq)+(unigram_freq[bigram[0]]-freq)))/num_reports  # d = N - (a + b + c) = num_reports-(freq)-(unigram_freq[bigram[1]]-freq)-(unigram_freq[bigram[0]]-freq)

            # PMI per bigram
            test_if_greater_than_0 = (((unigram_freq[bigram[0]] - freq) / num_reports) / (((freq + (unigram_freq[bigram[0]] - freq)) / num_reports) * ((freq + (unigram_freq[bigram[1]] - freq)) / num_reports)))
            if test_if_greater_than_0 >= 1:
                pmi_result['pmi_per_bigram'] = log(test_if_greater_than_0)
            pmi_result_list.append(pmi_result)
    return pmi_result_list


def get_chisquared_result(unigram_freq, bigram_freq):
    chi2_result_list = list()
    for bigram, freq in bigram_freq.items():
        if bigram[0] in unigram_freq and bigram[1] in unigram_freq:
            # calculate the Chi-square statistic for this bigram
            # you will need to use the following variables:
            #
            # freq: frequency of this bigram
            # unigram_freq[bigram[0]]: frequency of the first word in the bigram
            # unigram_freq[bigram[1]]: frequency of the second word in the bigram
            # num_reports: total number of tweets in the data (i.e. transactions in the database)
            #
            chi2_result = {'chi2_per_bigram': 0.0, 'bigram': str, 'freq': 0.0, 'bigram[0]': str, 'unigram_freq[bigram[0]]': 0.0, 'bigram[1]': str, 'unigram_freq[bigram[1]]': 0.0, 'has_word_1_has_word_2': 0.000, 'no_word_1_has_word_2': 0.000, 'has_word_1_no_word_2':0.000,'no_word_1_no_word_2':0.000}
            chi2_result['bigram']=bigram
            chi2_result['freq']=freq
            chi2_result['bigram[0]']=bigram[0]
            chi2_result['unigram_freq[bigram[0]]']=unigram_freq[bigram[0]]
            chi2_result['bigram[1]']=bigram[1]
            chi2_result['unigram_freq[bigram[1]]']=unigram_freq[bigram[1]]


            chi2_result['has_word_1_has_word_2']=freq/num_reports  # a = freq
            chi2_result['no_word_1_has_word_2']=(unigram_freq[bigram[0]]-freq)/num_reports  # b = unigram_freq[bigram[1]]-freq
            chi2_result['has_word_1_no_word_2']=(unigram_freq[bigram[1]]-freq)/num_reports   # c = unigram_freq[bigram[0]]-freq
            chi2_result['no_word_1_no_word_2']=(num_reports-((freq)+(unigram_freq[bigram[1]]-freq)+(unigram_freq[bigram[0]]-freq)))/num_reports  # d = N - (a + b + c) = num_reports-(freq)-(unigram_freq[bigram[1]]-freq)-(unigram_freq[bigram[0]]-freq)

            # PMI per bigram

            chi2_result['chi2_per_bigram']=((freq-((unigram_freq[bigram[0]]*unigram_freq[bigram[1]])/num_reports))**2)/((unigram_freq[bigram[0]]*unigram_freq[bigram[1]])/num_reports)
            chi2_result_list.append(chi2_result)
    return chi2_result_list


def unigram_to_excel(unigram_dict_to_excel, file_suffix):
    unigram_freq_2 = pd.DataFrame.from_dict(unigram_dict_to_excel, orient='index').reset_index()
    writer_c_2 = pd.ExcelWriter(current_directory + "\\Output\\" + file_suffix + "_unigram_freq.xlsx", engine='xlsxwriter')
    unigram_freq_2.to_excel(writer_c_2, sheet_name='main', index=False)
    writer_c_2.save()
    print("Unigram frequency complete and '" + file_suffix + "_unigram_freq.xlsx' is saved to the Output folder.")


def bigram_to_excel(bigram_dict_to_excel, file_suffix):
    bigram_freq_2 = pd.DataFrame.from_dict(bigram_freq, orient='index').reset_index()
    writer_bigram_word_list = pd.ExcelWriter(current_directory + "\\Output\\" + file_suffix + "_bigram_freq.xlsx", engine='xlsxwriter')
    bigram_freq_2.to_excel(writer_bigram_word_list, sheet_name='main')
    writer_bigram_word_list.save()
    print("Bigram frequency complete and '" + file_suffix + "_bigram_freq.xlsx is saved to the Output folder.")


print()

# # Get unigram freq and save to excel
unigram_freq = get_itemset_frequency(preprocessed_desc_df, process_text_into_word_list)
unigram_to_excel(unigram_freq, 'all')

# # Get bigram freq and export to excel
bigram_freq = get_itemset_frequency(preprocessed_desc_df, process_text_into_bigram_list)
bigram_to_excel(bigram_freq, 'all')

# To print unigram or bigram, use 'print(bigram_freq)' or:
# for bigram, freq in sorted(bigram_freq.items(), key = lambda x: x[1], reverse = True)[:100]:

# # Generate PMI results and export to excel
all_pmi_result_list = get_pmi_result(unigram_freq, bigram_freq)
all_pmi_results = pd.DataFrame(all_pmi_result_list)
# print(all_pmi_results)

# # Generate chi2 results and export to excel
chi2_df = get_chisquared_result(unigram_freq, bigram_freq)
all_chi2_results = pd.DataFrame(chi2_df)
# print(all_chi2_results)

# Export PMI and Chi-squared results as one file, print file stats
all_pmi_results.insert(1, 'chi2_per_bigram', all_chi2_results['chi2_per_bigram'])
all_pmi_chisquared_results = all_pmi_results.copy()
export_to_excel(all_pmi_chisquared_results)
print()
print("See PMI and chi squared result for bigrams, below:")
print(all_pmi_chisquared_results.describe())
print(all_pmi_chisquared_results)


# 4 # Identify clusters using semi-processed text and print unigrams, bigrams, PMI, and chi-squared results
print()

from sklearn.feature_extraction.text import TfidfVectorizer
vectorizer = TfidfVectorizer()
X = vectorizer.fit_transform(df['Description_stemmer'])
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
plt.savefig(current_directory + "\\Output\\" + "sum_of_squared_distance.jpg", format="jpg")
# plt.show()
print("Sum of squares plot complete. See 'sum_of_squared_distance.pdf'")

true_k = 3 # manually set
model = KMeans(n_clusters=true_k, init='k-means++', max_iter=200, n_init=10)
model.fit(X)
labels=model.labels_
report_clusters=pd.DataFrame(list(zip(df['Description_stemmer'], labels)), columns=['title', 'cluster']).join(df)

# # Save clusters to excel
export_to_excel(report_clusters)
print("Clustering complete.")

# # Create a dataframes for each cluster and print stats
for x in range(0, true_k):
    d = {}
    indiv_cluster_name = [x for x in globals() if globals()[x] is report_clusters][0] + "_" + str(x)
    d[indiv_cluster_name] = report_clusters.loc[(report_clusters['cluster'] == x)]
    # print(indiv_cluster_name)

    # # Define variables
    cluster_iter = x + 1
    file_suffix = str(cluster_iter) + "of" + str(true_k)

    # # Print column stats for each cluster
    print()
    cluster_df = pd.DataFrame(d[indiv_cluster_name]).copy()
    cluster_df['Created At'] = pd.to_datetime(cluster_df['Created At'])
    # print(cluster_df)
    num_reports = print_column_stats(cluster_df, file_suffix + ' cluster')

    cluster_desc_df = cluster_df["Description_stemmer"]
    # # Get unigram freq and save to excel
    cluster_unigram_freq = get_itemset_frequency(cluster_df, process_text_into_word_list)
    unigram_to_excel(cluster_unigram_freq, file_suffix)

    # # Get bigram freq and export to excel
    cluster_bigram_freq = get_itemset_frequency(cluster_df, process_text_into_bigram_list)
    bigram_to_excel(cluster_bigram_freq, file_suffix)

    # To print unigram or bigram, use 'print(d[file_suffix + "_bigram_freq"])' or:
    # for bigram, freq in sorted(bigram_freq.items(), key = lambda x: x[1], reverse = True)[:100]:

    # # Generate PMI results and export to excel
    pmi_df = get_pmi_result(unigram_freq, bigram_freq)
    pmi_result_df = pd.DataFrame(pmi_df)
    # print(all_pmi_results)

    # # Generate chi2 results and export to excel
    chisquared_df = get_chisquared_result(unigram_freq, bigram_freq)
    chi2_result_df = pd.DataFrame(chisquared_df)
    # print(chi2_df)

    # Export PMI and Chi-squared results as one file, print file stats
    pmi_result_df.insert(1, 'chi2_per_bigram', chi2_result_df['chi2_per_bigram'])
    pmi_chisquared_results = pmi_result_df.copy()
    dict_export_to_excel(pmi_chisquared_results, file_suffix + "_pmi_chi_squared_results")
    print()
    print("See PMI and chi squared result for " + file_suffix + "cluster bigrams, below:")
    print(pmi_chisquared_results.describe())
    print(pmi_chisquared_results)

print("Cluster text analysis complete.")

# 5 # Print word clouds for each cluster
result = {'cluster': labels, 'wiki': report_clusters['Description_stemmer']}
result = pd.DataFrame(result)

stopwords = set(STOPWORDS)
stopwords.update(["patient", "pt", "patients"])

for k in range(0, true_k):
   s = result[result.cluster==k]
   k_plus1 = k + 1
   print("There are {} words in the " + str(k) + " cluster.".format(len(s['wiki'])))
   text = s['wiki'].str.cat(sep=' ')
   text = text.lower()
   text = ' '.join([word for word in text.split()])
   wordcloud = WordCloud(stopwords=stopwords, max_font_size=50, max_words=100, background_color="white").generate(text)
   # print('Cluster: {}'.format(k))
   # print('Titles')
   titles = report_clusters[report_clusters.cluster==k]['title']
   # print(titles.to_string(index=False))
   plt.figure()
   plt.imshow(wordcloud, interpolation="bilinear")
   plt.axis("off")
   plt.savefig(current_directory + "\\Output\\" + "cluster_" + str(k_plus1) + "_of_" + str(true_k) + "_word_cloud.jpg", format="jpg")
#  plt.show()

print("Word clouds for clusters complete.")


# 6 # Create a dataframe for pre-covid data and print stats and word cloud
report_clusters['Created At'] = pd.to_datetime(report_clusters['Created At'])

file_suffix = 'pre_Covid'

# Create Pre-Covid dataframe and describe
pre_covid_df = report_clusters.loc[(report_clusters['Created At'] <= '03-31-2020')]
pre_covid_df['Created At'] = pd.to_datetime(pre_covid_df['Created At'])
num_reports = print_column_stats(pre_covid_df, file_suffix)
print(pre_covid_df.describe())
# print(pre_covid_df)

# Add description column to its own dataframe
pre_covid_desc_df = pre_covid_df['Description_stemmer'].copy()

# # Get unigram freq and save to excel
precovid_unigram_freq = get_itemset_frequency(pre_covid_desc_df, process_text_into_word_list)
unigram_to_excel(precovid_unigram_freq, file_suffix)

# # Get bigram freq and export to excel
precovid_bigram_freq = get_itemset_frequency(pre_covid_desc_df, process_text_into_bigram_list)
bigram_to_excel(precovid_bigram_freq, file_suffix)

# To print unigram or bigram, use 'print(d[file_suffix + "_bigram_freq"])' or:
# for bigram, freq in sorted(bigram_freq.items(), key = lambda x: x[1], reverse = True)[:100]:

# # Generate PMI results and export to excel
pmi_df = get_pmi_result(precovid_unigram_freq, precovid_bigram_freq)
pmi_result_df = pd.DataFrame(pmi_df)
# print(all_pmi_results)

# # Generate chi2 results and export to excel
chisquared_df = get_chisquared_result(precovid_unigram_freq, precovid_bigram_freq)
chi2_result_df = pd.DataFrame(chisquared_df)
# print(chi2_df)

# Export PMI and Chi-squared results as one file, print file stats
pmi_result_df.insert(1, 'chi2_per_bigram', chi2_result_df['chi2_per_bigram'])
precovid_pmi_chisquared_results = pmi_result_df.copy()
export_to_excel(precovid_pmi_chisquared_results)
print()
print("See PMI and chi squared result for pre-covid bigrams, below:")
print(precovid_pmi_chisquared_results.describe())
print(precovid_pmi_chisquared_results)

print("Pre-covid text analysis complete.")


# # Print word clouds for each cluster
text = " ".join(review for review in pre_covid_df["Description_stemmer"])
print("There are {} words in the post-covid data set.".format(len(text)))

stopwords = set(STOPWORDS)
stopwords.update(["patient", "pt", "patients"])

# Generate a word cloud image
wordcloud = WordCloud(stopwords=stopwords, max_font_size=50, max_words=100, background_color="white").generate(text)
# print('Cluster: {}'.format(k))
# print('Titles')
# titles = report_clusters['title']
# print(titles.to_string(index=False))
plt.figure()
plt.imshow(wordcloud, interpolation="bilinear")
plt.axis("off")
plt.savefig(current_directory + "\\Output\\" + "pre_covid" + "_word_cloud.jpg", format="jpg")
#  plt.show()

print("Pre-covid word clouds complete.")



# 6 # Create a dataframes for each cluster and print stats and word cloud
report_clusters['Created At'] = pd.to_datetime(report_clusters['Created At'])

file_suffix = 'post_Covid'

# Create Post-Covid dataframe and describe
post_covid_df = report_clusters.loc[(report_clusters['Created At'] >= '04-01-2020')]
post_covid_df['Created At'] = pd.to_datetime(post_covid_df['Created At'])
num_reports = print_column_stats(post_covid_df, file_suffix)
print(post_covid_df.describe())
# print(post_covid_df)

# Add description column to its own dataframe
post_covid_desc_df = post_covid_df['Description_stemmer'].copy()

# # Get unigram freq and save to excel
postcovid_unigram_freq = get_itemset_frequency(post_covid_desc_df, process_text_into_word_list)
unigram_to_excel(precovid_unigram_freq, file_suffix)

# # Get bigram freq and export to excel
postcovid_bigram_freq = get_itemset_frequency(post_covid_desc_df, process_text_into_bigram_list)
bigram_to_excel(postcovid_bigram_freq, file_suffix)

# To print unigram or bigram, use 'print(d[file_suffix + "_bigram_freq"])' or:
# for bigram, freq in sorted(bigram_freq.items(), key = lambda x: x[1], reverse = True)[:100]:

# # Generate PMI results and export to excel
pmi_df = get_pmi_result(postcovid_unigram_freq, postcovid_bigram_freq)
pmi_result_df = pd.DataFrame(pmi_df)
# print(pmi_result_df)

# # Generate chi2 results and export to excel
chisquared_df = get_chisquared_result(postcovid_unigram_freq, postcovid_bigram_freq)
chi2_result_df = pd.DataFrame(chisquared_df)
# print(chi2_result_df)

# Export PMI and Chi-squared results as one file, print file stats
pmi_result_df.insert(1, 'chi2_per_bigram', chi2_result_df['chi2_per_bigram'])
postcovid_pmi_chisquared_results = pmi_result_df.copy()
export_to_excel(postcovid_pmi_chisquared_results)
print()
print("See PMI and chi squared result for post-covid bigrams, below:")
print(postcovid_pmi_chisquared_results.describe())
print(postcovid_pmi_chisquared_results)


print("Post-covid text analysis complete.")

# # Print word clouds for each cluster
text = " ".join(review for review in post_covid_df["Description_stemmer"])
print("There are {} words in the post-covid data set.".format(len(text)))

stopwords = set(STOPWORDS)
stopwords.update(["patient", "pt", "patients"])

# Generate a word cloud image
wordcloud = WordCloud(stopwords=stopwords, max_font_size=50, max_words=100, background_color="white").generate(text)
# print('Cluster: {}'.format(k))
# print('Titles')
# titles = report_clusters['title']
# print(titles.to_string(index=False))
plt.figure()
plt.imshow(wordcloud, interpolation="bilinear")
plt.axis("off")
plt.savefig(current_directory + "\\Output\\" + "post_covid" + "_word_cloud.jpg", format="jpg")
#  plt.show()

print("Post-covid word clouds complete.")




