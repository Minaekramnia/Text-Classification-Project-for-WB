# -*- coding: utf-8 -*-
"""
Created on Thu Aug 15 15:22:24 2019

@author: wb550776
"""

import nltk
from nltk.stem import PorterStemmer
from nltk.tokenize import sent_tokenize
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.model_selection import train_test_split
from sklearn.naive_bayes import MultinomialNB
# from sklearn.naive_bayes import ComplementNB
from sklearn.ensemble import RandomForestClassifier
from sklearn.neural_network import MLPClassifier
import numpy as np
import pandas as pd
import re
import pickle
import os
from docx import Document
from sklearn.metrics.pairwise import cosine_similarity
from matplotlib import pyplot as plt


##### PART 1 Use pre-trained model #####
# load and clean data

lesson = pd.read_excel(r'C:\Users\wb550776\Documents\Projects\lessons.xlsx')

tokenized_text=sent_tokenize(text)
tokenized_text=sent_tokenize("lesson['Future']")
lesson.query("ProjectId ==2116")


# load customized stopwords
stopwords = set(nltk.corpus.stopwords.words('english'))
#cust_stopwords = pd.read_csv(strPath + 'stopwords.csv', header = None)


stopwords.update(set(cust_stopwords[0]))


folderPath = os.path.join(os.getcwd(),'documents')
#Tokenization
## Tokenize
# Tokenize with stemming
def tokenize_and_stem(text):
    # first tokenize by sentence, then by word to ensure that punctuation is caught as it's own token
    tokens = [word for sent in nltk.sent_tokenize(text) for word in nltk.word_tokenize(sent)]
    filtered_tokens = []
    # filter out any tokens not containing letters (e.g., numeric tokens, raw punctuation)
    for token in tokens:
        token = re.sub('[^\w+]', "", token)
        if re.search('[a-zA-Z]', token):
            if token not in stopwords:
                filtered_tokens.append(token)
    stems = [PorterStemmer().stem(word) for word in filtered_tokens]
    return stems

# Tokenize without stemming
def tokenize(text):
    # first tokenize by sentence, then by word to ensure that punctuation is caught as it's own token
    tokens = [word for sent in nltk.sent_tokenize(text) for word in nltk.word_tokenize(sent)]
    filtered_tokens = []
    # filter out any tokens not containing letters (e.g., numeric tokens, raw punctuation)
    for token in tokens:
        token = re.sub('[^\w+]', "", token)
        if re.search('[a-zA-Z]', token):
            if token not in stopwords:
                filtered_tokens.append(token)
    return filtered_tokens


