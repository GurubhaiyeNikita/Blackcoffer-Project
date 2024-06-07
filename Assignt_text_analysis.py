import re
import pandas as p
import requests
from bs4 import BeautifulSoup
import nltk
import openpyxl
import xlsxwriter
from nltk.corpus import cmudict
from nltk import NLTKWordTokenizer
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from nltk.tokenize import word_tokenize, sent_tokenize
nltk.download('punkt')
nltk.download('cmudict')
syllable_dict = cmudict.dict()
nltk.download('stopwords')


#to create an output file with desired columns
col_name=("URLID","URL","Positive word count","Negative word count","Polarity","Subjectivity","Avg sentence Length","percent of complex words","Fog index","Avg number of words per sentence",'Complex word count','Total words',"Syllable count per word","Personal pronouns","Avg word length")



links=p.read_excel("/Input.xlsx")
l=len(links)


for i in range(0,l):

 url=links.loc[i]['URL']
 open=requests.get(url)
 htmlc=open.text
 code=BeautifulSoup(open.content,"html.parser")
 code.prettify()

 title=code.title.string
 #print(title)


 txt=code.find("div",{"class":["td-pb-span8 td-main-content","td-ss-main-content","td-pb-row","td-post-content tagdiv-type","td-container"]}).get_text()
#print(txt)
 text=title +" "+ txt           #adding title with the rest of article
 text_size=len(text)   # text variable can print entire article




#removing stopwords
 stopwords_default = stopwords.words('english')
#print('Stopwords in NLTK: ',len(stopwords_default))
 text_tokens = word_tokenize(text)
 stop_clean_text = [word for word in text_tokens if not word in stopwords_default]
 w_after_stop__clean=len(' '.join(stop_clean_text))


 #removing punctuations
 py_opstr = re.sub(r'[^\w\s]','',text)    #to remove punctuations
#print ('Without punctuation string: ', py_opstr)
 clean_text = [word for word in py_opstr if not word in stopwords_default]

 w_after_clean=len(' '.join(clean_text))

# Define simple keyword lists for positive, negative, and neutral sentiments
 positive_count=0
 negative_count=0
 f=p.read_excel("/Positive and Negative Word List.xlsx")

 lg=len(f)

 #print(text1)
 for g in range(0,lg):
  fp=f.loc[g]['Positive Sense Word List']
  fn=f.loc[g]['Negative Sense Word List']
  for word in text.split():
          if fp==word:
            positive_count+=1
          elif fn==word:
            negative_count+=1
          else:
            pass

 #print("\npositive score",positive_count)
 #print("\nnegative score",negative_count)

 polarity=(positive_count-negative_count)/(positive_count+negative_count+0.0000001)
 #print("polarity",polarity)


 subjectivity=(positive_count+negative_count)/(w_after_clean+0.0000001)
#print(subjectivity)


 words =len(word_tokenize(text))
 sentences = len(sent_tokenize(text))  #find sentences
 

#print("Tokenized Words:", words)
#print("Tokenized Sentences:", sentences)
 try:
  avg_sentence_length=words/sentences   # as given in text analysis file
 except ZeroDivisionError:
  avg_sentence_length=words

 def count_syllables(word):
    if word.lower() not in syllable_dict:    # search for lower case version of the word in dictionary
        return 0
    return [len(list(y for y in x if y[-1].isdigit())) for x in syllable_dict[word.lower()]][0]
                                               # return number of syllable

 def is_complex(word):
    syllable_count = count_syllables(word)
    return syllable_count > 2

 def count_complex_words(text):
    words = nltk.word_tokenize(text)
    num_complex_words = sum(is_complex(word) for word in words)
    return num_complex_words

 comp_word_c = count_complex_words(text)
#print("\ncomplex",compword)
 try:
  perct_complex_words=comp_word_c/words
 except ZeroDivisionError:
  perct_complex_words=comp_word_c/100

 try:

  fog_index=0.4*(avg_sentence_length/perct_complex_words)

 except ZeroDivisionError:
  fog_index=0


 try:
  avg_n_word_per_sentence=text_size/ sentences
 except:
  avg_n_word_per_sentence=text_size


 def count_syllables(word):
    word = word.lower()
    num_vowels = len([char for char in word if char in 'aeiou'])
    if word.endswith('es') or word.endswith('ed'):
        num_vowels -= 1           # do not include 'es' and 'ed' of found
    return num_vowels

 def count_syllables_per_word(text):
    words = nltk.word_tokenize(text)
    syllable_counts = [count_syllables(word) for word in words]
    return syllable_counts

 syl_count_p_word=count_syllables_per_word(text)
 #print("\nsyllable count",sylcpw)


 def count_personal_pronouns(text):
    pattern = r"\b(I|we|my|ours|us)\b"   # pattern to check if those words exists
    pattern = r"(?<!\bUS\b)" + pattern   # pattern should not include US instead of us
    matches = re.findall(pattern, text, flags=re.IGNORECASE)
    count = len(matches)
    return count

 personal_pronoun=count_personal_pronouns(text)
 #print("\npersonal pronoun",pp)

 def calculate_avg_word_length(text):
    words = text.split()
    total_characters = sum(len(word) for word in words)  # sum of the charachter
    #print(total_characters)
    num_words = len(words)
    #print(num_words)                              # total of words
    try:
     avg_word_length = total_characters / num_words
    except:
      avg_word_length = total_characters
    return avg_word_length

 avg_word_length=calculate_avg_word_length(text)
#print("\naverage word length",awl)

 url_id="blackassign00" + str(i+1)

 import numpy as n

 data=n.array([url_id,url,positive_count,negative_count,polarity,subjectivity,avg_sentence_length,perct_complex_words,fog_index,avg_n_word_per_sentence,comp_word_c,word,syl_count_p_word,personal_pronoun,avg_word_length])


 d=data.reshape(1,15)


 df=p.DataFrame(d,columns=col_name)

#creates an excel file for first iteration 
 if(i==0):
   df.to_excel("/Output file.xlsx")
 else:
  with p.ExcelWriter("/Output file.xlsx",mode="a",engine="openpyxl",if_sheet_exists="overlay") as writer:
    df.to_excel(writer, sheet_name="Sheet1",header=None, startrow=writer.sheets["Sheet1"].max_row,index=False)


 df.to_csv("/Output file.csv",header=col_name, sep='|', index=False, mode='a')











