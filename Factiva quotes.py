import time
import sys
import os
import pandas as pd
import numpy as np
import stanza
import openpyxl
import re
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer



# stanza.download('en')
nlp = stanza.Pipeline('en')

#dfa = pd.read_pickle(r'C:/Users/danwilde/Dropbox (Penn)/Dissertation/Factiva/filtered_63_05_26.pkl')
#dfa = dfa.reset_index(drop=False)
#dfa.rename(columns = {'index':'old index'}, inplace = True)


#dfa = pd.read_excel(r'C:/Users/danwilde/Dropbox (Penn)/Dissertation/Factiva/filtered_63_05_26.xlsx')
dfa = pd.read_excel(r'C:/Users/danwilde/Dropbox (Penn)/Dissertation/Factiva/issues_2021_05_27_2125_articles.xlsx')


t0 = time.time()

# set up the job parameters
#INFILE = sys.argv[1]
#STARTROW = sys.argv[2]
#ENDROW = sys.argv[3]
#OUTDIR = "results"
#if not os.path.exists(OUTDIR):
#    os.makedirs(OUTDIR)

# get the article slice
#dfa = pd.read_pickle(INFILE)[int(STARTROW): int(ENDROW)]

columns = ['sent', 'firm', 'person', 'last', 'role',  'former', 'quote', 'sentence', 'qtype', 'qsaid',
           'qpref', 'comment', 'AN', 'date', 'source', 'regions', 'subjects', 'misc',
           'industries', 'focfirm', 'by', 'comp', 'sentiment_stanza', 'sent_vader_neg',
           'sent_vader_neu', 'sent_vader_pos']

df = pd.DataFrame(columns=columns)


stext = r'said|n.t say|say|told|n.t tell\S+|\btell\S+|assert\S+|according\s+to|n.t comment\S+|comment\S+|quote\S+' \
        r'|describ\S+|n.t communicat\S+|communicated|communicates|articulat\S+|n.t divulg\S+|divulg\S+|noted|noting' \
        r'|recounted|suggested|\bexplained\b|\badded\b|acknowledged|\bstating\b|\bstated\b|\bprotested\b|\bfumed\b' \
        r'|\bexpress\S+|announc\S+|called|recounts|\basks\b|\basked\b'

l1 = r'executive\s+vice\s+president\s+of\s\S+|executive\s+vice-president\s+of\s\S+|' \
     r'senior\s+vice\s+president\b\s+of|\S+\s+executive\b\s+of\s+the\s+\S+|vice\s+president\s+and\s+general\s+\S+'

l2 = r'executive\s+vice\s+president|executive\s+vice-president|sr\.\s+vice\s+president\b\s+of\s+\S|' \
     r'head\s+of\s\S+|chair\S+\s+and\s+\S+\s+head\b|senior\s+vice-president\b|' \
     r'senior\s+director\s+of\s+\S+|chief.+officer|president\s+and\s+C.0|chair\S+\s+and\s+\S+'

l3 = r'vice\s+chair\S+|vice\s+president|vice-president|senior\s+manager|sr.*manager|senior\s+director|' \
     r'chief\s+executive|senior\s+manager|sr.*manager|vp\s+of\s+\S+|global\+director|senior\s+director|' \
     r'sr.*director|executive\s+director|managing\s+director|vice\s+minister|prime\s+minister|founding\s+partner|board\s+member' \
     r'business\s+owner|sales\s+manager|svp|chief\s+engineer|chief\s+of\s+\S+|founder\s+and+S+'

l4 = r'president|chair\S+|director|manager|vp|' \
     r'banker\b|specialist\b|accountant\b|journalist\b|reporter\b|analyst\b|consultant\b|negotiator\b|' \
     r'governor\b|congress\S+|house\s+speaker|senator|senior\s+fellow|fellow|undersecretary|' \
     r'spokesman\b|spokesperson\b|spokeswoman\b|representative\b|official\b|executive\b|\baide\b|\bpilot\b|professor|' \
     r'attorney|minister|secretary|\bguru\b|partner|general\s+counsel|mayor|dealer|board|administrator|owner|' \
     r'leader|police\s+officer|police\s+chief|actor|actress|academy\s+award-winner|founder'

l5 = r'\bC.O\b'

leader_split = r'executive|president|director|manager|'\
            r'Executive|President|Director|Manager'
stext_split = r'said|told|comment'

pron = r'\bhe\b|\bshe\b'
plurals = r'analysts|bankers|consultants|executives|management|participants|officials|engineers'
unnamed = r'declined\s+to\s+be\s+named|unnamed|not named'
rep = r'spokesman\b|spokesperson\b|spokeswoman\b|representative\b|official\b|executive\b'
start = r'\"|\'\'|\`\`'
end = r'\"|\'\''
qs = r'\`\`|\"|\'\''
suf = r'\bJr\b|\bSr\b|\sI\s|\sII\s|\sIII\s|\bPhD\b|Ph\.D|\bMBA\b|\bCPA\b'
analyst = r'analyst|negotiator|\baide\b|\bpilot\b|professor|' \
     r'attorney|minister|secretary|\bguru\b|partner|general\s+counsel|mayor|dealer|board|administrator|owner|' \
     r'leader|president|police\s+officer|police\s+chief|academy\s+award-winner|actor|actress'
car_types = r'pickup\s+truck|truck|crossover|sedan|minivan|van|sports-utility\s+vehicle|sports\s+utility\s+vehicle|' \
            r'SUV|convertible\b|hatchback|station\s+wagon|coupe|sports|s+car|compact\s+car|roadster|mid-size\s+car|' \
            r'supercar|brand'
car_makers_names = r'Aston\s+Martin'
media = r'Reuters|Bloomberg|CNN|The\s+Independent|USA\s+Today|' \
        r'Press|Daily|Times|Chronicle|' \
        r'Journal|News|Herald|Post|Weekly|Telegraph|Wire|Bulletin|Inquirer|Digest'
prepositions =r'with|against|behind|before|because\s+of|following|from|in\s+spite\s+of|instead\s+of|' \
              r'than|over|under|versus|without|like|from'

maxch = str(200)
maxch1 = str(60)


# Set the max number of words away from the focal leader and the focal organization
mxld = 6
mxorg = 6
mxlds = 2
mxpron = 2
mxorg1 = 2
mxfrmer = 2
mxcar = 2
mxcomp = 100



# Create a class in which to place variables and then create a function that will place all the variables
# in a df row. Makes the code a bit more readable when you go through the todf function doezens of times
class Features:
    def __init__(self):
        self.df = self.qtype = self.qsaid = self.qpref = self.comment = self.n= self.person = \
        self.last = self.role = self.firm = self.former = self.quote = self.sentence = self.AN = self.date = self.source = \
        self.regions = self.subjects = self.misc = self.industries = self.focfirm = self.by = \
        self.comp = self.sentiment_stanza = self.sent_vader_neg = self.sent_vader_neu = self.sent_vader_pos = None




    def todf(self):
        df.loc[len(df.index)] = [self.n, self.firm, self.person, self.last, self.role,  self.former, self.quote,
                                           self.sentence, self.qtype, self.qsaid, self.qpref, self.comment, self.AN,
                                           self.date, self.source, self.regions, self.subjects, self.misc,
                                           self.industries, self.focfirm, self.by, self.comp, self.sentiment_stanza,
                                           self.sent_vader_neg, self.sent_vader_neu, self.sent_vader_pos]

        dfsent.at[n - 1, 'covered'] = 1

    def todf_info(self):
        df.loc[len(df.index)] = [self.n, self.firm, self.person, self.last, self.role,  self.former, self.quote,
                                           self.sentence, self.qtype, self.qsaid, self.qpref, self.comment, self.AN,
                                           self.date, self.source, self.regions, self.subjects, self.misc,
                                           self.industries, self.focfirm, self.by, self.comp, self.sentiment_stanza,
                                           self.sent_vader_neg, self.sent_vader_neu, self.sent_vader_pos]



#Initation this class
h = Features()




#Function to produce multiple sentence quotes
def multisent(stence):
    # check if a quote begins in this sentence:
    msent_quotes = []
    msent_quote = re.findall(r'(' + start + r')(.+)', stence, re.IGNORECASE)
    msent_quote = ''.join(''.join(elems) for elems in msent_quote)
    msent_quote = re.sub(r'(' + qs + r')', '', msent_quote)
    msent_quotes.append(msent_quote)
    stences = []
    stences.append(stence)
    # Create list of sentiments
    sentiments = []
    sent_vader_negs = []
    sent_vader_neus = []
    sent_vader_poss = []
    # Sentiment of current sentence, which is index n-1
    k = doc.sentences[n-1]
    sentiment_stanza = k.sentiment
    sent_vader_neg = sentiment_scores(k.text)[0]
    sent_vader_neu = sentiment_scores(k.text)[1]
    sent_vader_pos = sentiment_scores(k.text)[2]
    sent_vader_negs.append(sent_vader_neg)
    sent_vader_neus.append(sent_vader_neu)
    sent_vader_poss.append(sent_vader_pos)
    sentiments.append(sentiment_stanza)
    first_before_quote = re.findall(r'(.+)(' + start + r')', stence, re.IGNORECASE)
    first_before_quote = ''.join(''.join(elems) for elems in first_before_quote)
    first_before_quote = re.sub(r'(' + qs + r')', '', first_before_quote)
    first_said = ""
    first_pron = ""
    first_ppl = ""
    last_after_quote = ""
    last_sindex = 0
    last_said = ""
    last_pron = ""
    last_ppl = ""
    nindex = []

    # If a quote does begin in this sentence, concatinate the subsequent sentences until the quote ends
    if len(msent_quote) > 0:
        nindex.append(n)
        for sindex in range(n, len(doc.sentences)):
            dfsent.at[sindex - 1, 'covered'] = 1
            j = doc.sentences[sindex]
            msent = j.text
            sentiment_stanza = j.sentiment
            sentiments.append(sentiment_stanza)
            sent_vader_neg = sentiment_scores(msent)[0]
            sent_vader_neu = sentiment_scores(msent)[1]
            sent_vader_pos = sentiment_scores(msent)[2]
            sent_vader_negs.append(sent_vader_neg)
            sent_vader_neus.append(sent_vader_neu)
            sent_vader_poss.append(sent_vader_pos)
            lastmsent_quote = re.findall(r'(.+)(' + end + r')', msent, re.IGNORECASE)
            lastmsent_quote = ''.join(''.join(elems) for elems in lastmsent_quote)
            lastmsent_quote = re.sub(r'(' + qs + r')', '', lastmsent_quote)
            stences.append(msent)
            # If there's not ending of a multi-sentence quote, keep appending
            if len(lastmsent_quote) == 0:
                msent_quote = msent_quote + " " + msent
                msent_quotes.append(msent)
                dfsent.at[sindex, 'covered'] = 1
                nindex.append(sindex+1)
            # If there is an ending of a multi-sentence quote, append and you're done.
            elif len(lastmsent_quote) > 0:
                msent_quote = msent_quote + " " + lastmsent_quote
                last_after_quote = re.findall(r'(' + end + r')(.+)', msent, re.IGNORECASE)
                last_after_quote = ''.join(''.join(elems) for elems in last_after_quote)
                msent_quotes.append(lastmsent_quote)
                last_sindex = sindex
                dfsent.at[sindex, 'covered'] = 1
                nindex.append(sindex+1)
                break

        first_said = re.findall(r'' + stext + r'', first_before_quote, re.IGNORECASE)
        first_pron = re.findall(r'' + pron + r'', first_before_quote, re.IGNORECASE)
        first_ppl = dfsent["ppl"][n-1]
        last_said = re.findall(r'' + stext + r'', last_after_quote, re.IGNORECASE)
        last_pron = re.findall(r'' + pron + r'', last_after_quote, re.IGNORECASE)
        last_ppl = dfsent["ppl"][last_sindex]

    # the output from the function
    return last_sindex, msent_quote, last_after_quote, first_said, first_pron, first_ppl, \
           last_said, last_pron, last_ppl, msent_quotes, stences, sentiments, nindex, sent_vader_negs, \
           sent_vader_neus, sent_vader_poss



def quote_set_even(sentence):
    quotes = []
    # split the sentence by quotation mark
    quote_split = re.split(r'(' + qs + r')', sentence)
    # Remove blanks if any
    quote_split = list(filter(None, quote_split))
    # Remove any items that are quotation marks
    rex = re.compile(r'' + qs + r'')
    quote_split = [x for x in quote_split if not rex.match(x)]

    for q in range(len(quote_split)):
        quote_beginning = re.findall(r'^\s*(' + start + r')', stence, re.IGNORECASE)
        if len(quote_beginning) > 0:
            # Want position 1, 3, etc.
            if q % 2 == 0:
                quotes.append(quote_split[q])
        if len(quote_beginning) == 0:
            # Want position 0, 2, 4, etc.
            if q % 2 == 1:
                quotes.append(quote_split[q])
    return quotes

def quote_set(sentence):
    quotes =[]
    last_sindex = msent_quote = last_after_quote = first_said = \
        first_pron = first_ppl = last_said = last_pron = last_ppl = msent_quotes = stences = \
        sentiments = nindex = sent_vader_negs =  sent_vader_neus = sent_vader_poss = None
    multi_sent_indicator = 0
    # Since no clear person reference, check for FULL quotes:
    q_marks = re.findall(r'(' + qs + r')', sentence, re.IGNORECASE)
    # If odd number then capture all quotes and also start multiquote on last one.

    # If there are an even number of quotation marks
    if len(q_marks) % 2 == 0:
        quotes = quote_set_even(sentence)


    elif len(q_marks) % 2 == 1:
        multi_sent_indicator = 1
        # multisentence for last one
        if len(q_marks) == 1:
            last_sindex = multisent(sentence)[0]
            msent_quote = multisent(sentence)[1]
            last_after_quote = multisent(sentence)[2]
            first_said = multisent(sentence)[3]
            first_pron = multisent(sentence)[4]
            first_ppl = multisent(sentence)[5]
            last_said = multisent(sentence)[6]
            last_pron = multisent(sentence)[7]
            last_ppl = multisent(sentence)[8]
            msent_quotes = multisent(sentence)[9]
            stences = multisent(sentence)[10]
            sentiments = multisent(sentence)[11]
            nindex = multisent(sentence)[12]
            sent_vader_negs = multisent(sentence)[13]
            sent_vader_neus = multisent(sentence)[14]
            sent_vader_poss = multisent(sentence)[15]
        else:
            # Drop last quote to capture any full quotes
            stence_drop_last_quote = re.sub('(.*)(' + qs + ')(.*)', r'\1 \3', stence)
            quotes = quote_set_even(stence_drop_last_quote)
            # Drop all quotation marks but last one and plug into multisent function
            all_but_last = re.sub(r'(' + qs + ')', '', stence, count=len(q_marks) - 1)
            last_sindex = multisent(all_but_last)[0]
            msent_quote = multisent(all_but_last)[1]
            last_after_quote = multisent(all_but_last)[2]
            first_said = multisent(all_but_last)[3]
            first_pron = multisent(all_but_last)[4]
            first_ppl = multisent(all_but_last)[5]
            last_said = multisent(all_but_last)[6]
            last_pron = multisent(all_but_last)[7]
            last_ppl = multisent(all_but_last)[8]
            msent_quotes = multisent(all_but_last)[9]
            stences = multisent(all_but_last)[10]
            sentiments = multisent(all_but_last)[11]
            nindex = multisent(all_but_last)[12]
            sent_vader_negs  = multisent(all_but_last)[13]
            sent_vader_neus = multisent(all_but_last)[14]
            sent_vader_poss = multisent(all_but_last)[15]


    return last_sindex, msent_quote, last_after_quote, first_said, first_pron, first_ppl, \
           last_said, last_pron, last_ppl, msent_quotes, stences, sentiments, nindex, quotes, multi_sent_indicator, \
           sent_vader_negs, sent_vader_neus, sent_vader_poss

# Function for when you want to search with instance of a role and another variable, when distance between
# doesn't matter.
def longrolesearch(sentence, non_role_var):
    role_list1a = re.findall(r'(' + non_role_var + r').+(' + l1 + r')', sentence, re.IGNORECASE)
    role_list1b = re.findall(r'(' + non_role_var + r').+(' + l2 + r')', sentence, re.IGNORECASE)
    role_list1c = re.findall(r'(' + non_role_var + r').+(' + l3 + r')', sentence, re.IGNORECASE)
    role_list1d = re.findall(r'(' + non_role_var + r').+(' + l4 + r')', sentence, re.IGNORECASE)
    role_list1e = re.findall(r'(' + non_role_var + r').+(' + l5 + r')', sentence)
    role_list2a = re.findall(r'(' + l1 + r').+(' + non_role_var + r')', sentence, re.IGNORECASE)
    role_list2b = re.findall(r'(' + l2 + r').+(' + non_role_var + r')', sentence, re.IGNORECASE)
    role_list2c = re.findall(r'(' + l3 + r').+(' + non_role_var + r')', sentence, re.IGNORECASE)
    role_list2d = re.findall(r'(' + l4 + r').+(' + non_role_var + r')', sentence, re.IGNORECASE)
    role_list2e = re.findall(r'(' + l5 + r').+(' + non_role_var + r')', sentence)

    role_list = role_list1a + role_list1b + role_list1c + role_list1d + role_list1e + role_list2a + \
                role_list2b + role_list2c + role_list2d + role_list2e

    return role_list

# Function for when you want to search with instance of a role and another variable, when distance between
# them matters.
def shortrolesearch(sentence, non_role_var, max_char):
    role_list1a = re.findall(
        r'\b(?:(' + non_role_var + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?(' + l1 + r'))',
        sentence,
        re.IGNORECASE)
    role_list1b = re.findall(
        r'\b(?:(' + non_role_var + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?(' + l2 + r'))',
        sentence,
        re.IGNORECASE)
    role_list1c = re.findall(
        r'\b(?:(' + non_role_var + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?(' + l3 + r'))',
        sentence,
        re.IGNORECASE)
    role_list1d = re.findall(
        r'\b(?:(' + non_role_var + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?(' + l4 + r'))',
        sentence,
        re.IGNORECASE)
    role_list1e = re.findall(
        r'\b(?:(' + non_role_var + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?(' + l5 + r'))',
        sentence)
    role_list2a = re.findall(
        r'\b(?:(' + l1 + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?(' + non_role_var + r'))',
        sentence,
        re.IGNORECASE)
    role_list2b = re.findall(
        r'\b(?:(' + l2 + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?(' + non_role_var + r'))',
        sentence,
        re.IGNORECASE)
    role_list2c = re.findall(
        r'\b(?:(' + l3 + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?(' + non_role_var + r'))',
        sentence,
        re.IGNORECASE)
    role_list2d = re.findall(
        r'\b(?:(' + l4 + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?(' + non_role_var + r'))',
        sentence,
        re.IGNORECASE)
    role_list2e = re.findall(
        r'\b(?:(' + l5 + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?(' + non_role_var + r'))',
        sentence)

    role_list = role_list1a + role_list1b + role_list1c + role_list1d + role_list1e + role_list2a + \
                role_list2b + role_list2c + role_list2d + role_list2e
    return role_list



def whoserolesearch(sentence,non_role_var, max_char):
    o3a = re.findall(
        r'\b(?:(' + non_role_var + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?whose\s+(' + l1 + '))',
        sentence,
        re.IGNORECASE)
    o3b = re.findall(
        r'\b(?:(' + non_role_var + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?whose\s+(' + l2 + '))',
        sentence,
        re.IGNORECASE)
    o3c = re.findall(
        r'\b(?:(' + non_role_var + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?whose\s+(' + l3 + '))',
        sentence,
        re.IGNORECASE)
    o3d = re.findall(
        r'\b(?:(' + non_role_var + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?whose\s+(' + l4 + '))',
        sentence,
        re.IGNORECASE)
    o3e = re.findall(
        r'\b(?:(' + non_role_var + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?whose\s+(' + l5 + '))',
        sentence)

    org_find = o3a + o3b + o3c + o3d + o3e

    return org_find

def ofrolesearch(sentence, non_role_var, max_char):
    z1 = re.findall(
        r'\b(?:(' + l1 + r')(\s+of)\W+(?:\w+\W+){0,' + str(max_char) + r'}?(' + non_role_var + '))', sentence,
        re.IGNORECASE)
    z2 = re.findall(
        r'\b(?:(' + l2 + r')(\s+of)\W+(?:\w+\W+){0,' + str(max_char) + r'}?(' + non_role_var + '))', sentence,
        re.IGNORECASE)
    z3 = re.findall(
        r'\b(?:(' + l3 + r')(\s+of)\W+(?:\w+\W+){0,' + str(max_char) + r'}?(' + non_role_var + '))', sentence,
        re.IGNORECASE)
    z4 = re.findall(
        r'\b(?:(' + l4 + r')(\s+of)\W+(?:\w+\W+){0,' + str(max_char) + r'}?(' + non_role_var + '))', sentence,
        re.IGNORECASE)
    z5 = re.findall(
        r'\b(?:(' + l5 + r')(\s+of)\W+(?:\w+\W+){0,' + str(max_char) + r'}?(' + non_role_var + '))', sentence)

    org_find = z1 + z2 + z3 + z4 + z5
    return org_find


def companysrolesearch(sentence,non_role_var, max_char):
    o3a = re.findall(
        r'\b(?:(' + non_role_var + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?company\'s\s+(' + l1 + '))',
        sentence,
        re.IGNORECASE)
    o3b = re.findall(
        r'\b(?:(' + non_role_var + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?company\'s\s+(' + l2 + '))',
        sentence,
        re.IGNORECASE)
    o3c = re.findall(
        r'\b(?:(' + non_role_var + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?company\'s\s+(' + l3 + '))',
        sentence,
        re.IGNORECASE)
    o3d = re.findall(
        r'\b(?:(' + non_role_var + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?company\'s\s+(' + l4 + '))',
        sentence,
        re.IGNORECASE)
    o3e = re.findall(
        r'\b(?:(' + non_role_var + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?company\'s\s+(' + l5 + '))',
        sentence)

    org_find = o3a + o3b + o3c + o3d + o3e

    return org_find

def threerolesearch(sentence, non_role_var1, non_role_var2):
    c1 = re.findall(r'' + non_role_var1 + r'.+(' + l1 + r').+(' + non_role_var2 + r')', sentence,
                    re.IGNORECASE)
    c2 = re.findall(r'' + non_role_var1 + r'.+(' + l2 + r').+(' + non_role_var2 + r')', sentence,
                    re.IGNORECASE)
    c3 = re.findall(r'' + non_role_var1 + r'.+(' + l3 + r').+(' + non_role_var2 + r')', sentence,
                    re.IGNORECASE)
    c4 = re.findall(r'' + non_role_var1 + r'.+(' + l4 + r').+(' + non_role_var2 + r')', sentence,
                    re.IGNORECASE)
    c5 = re.findall(r'' + non_role_var1 + r'.+(' + l5 + r').+(' + non_role_var2 + r')', sentence)
    c6 = re.findall(r'' + non_role_var1 + r'.+(' + non_role_var2 + r').+(' + l1 + r')', sentence,
                    re.IGNORECASE)
    c7 = re.findall(r'' + non_role_var1 + r'.+(' + non_role_var2 + r').+(' + l2 + r')', sentence,
                    re.IGNORECASE)
    c8 = re.findall(r'' + non_role_var1 + r'.+(' + non_role_var2 + r').+(' + l3 + r')', sentence,
                    re.IGNORECASE)
    c9 = re.findall(r'' + non_role_var1 + r'.+(' + non_role_var2 + r').+(' + l4 + r')', sentence,
                    re.IGNORECASE)
    c10 = re.findall(r'' + non_role_var1 + r'.+(' + non_role_var2 + r').+(' + l5 + r')', sentence)

    org_find = c1 + c2 + c3 + c4 + c5 + c6 + c7 + c8 + c9 + c10


    return org_find

def rolesearch(sentence):
        p1a = re.findall(r'(' + l1 + r')', sentence, re.IGNORECASE)
        p1b = re.findall(r'(' + l2 + r')', sentence, re.IGNORECASE)
        p1c = re.findall(r'(' + l3 + r')', sentence, re.IGNORECASE)
        p1d = re.findall(r'(' + l4 + r')', sentence, re.IGNORECASE)
        p1e = re.findall(r'(' + l5 + r')', sentence)
        role_list = p1a + p1b + p1c + p1d + p1e

        return role_list



def info_extraction(stence, orgs):
    # Add "Information" regarding any company or person to the main database.
    # First, if just one org, add that information
    if len(orgs) == 1:
        h.person = "NA"
        h.last = "NA"
        h.firm = orgs[0]
        h.role = "NA"
        h.former = "NA"
        h.qtype = "information"
        h.qsaid = "NA"
        h.qpref = "NA"
        h.comment = "information about one firm"
        h.quote = h.sentence = stence
        h.todf_info()

    # Second, if multiple orgs, add that information
    else:
        for o in orgs:
            h.person = "NA"
            h.last = "NA"
            h.firm = o
            h.role = "NA"
            h.former = "NA"
            h.qtype = "information"
            h.qsaid = "NA"
            h.qpref = "NA"
            h.comment = "information about multiple firm"
            h.quote = h.sentence = stence
            h.todf_info()


    # Next, capture any data on person
    for index, per in dfpeople.iterrows():
        last = per['last']
        people_string = re.findall(r'(' + last + r')', stence, re.IGNORECASE)
        if len(people_string) > 0:
            h.person = per['full']
            h.last = per['last']
            h.firm = per['firm']
            h.role = per['role']
            h.former = per['former']
            h.qtype = "information"
            h.qsaid = "NA"
            h.qpref = "NA"
            h.comment = "information about person"
            h.quote = h.sentence = stence
            h.todf_info()

def role_set(role_list, var_to_strip):
    role = ""
    for i in role_list:
        j = ''.join(''.join(elems) for elems in i)
        role = re.sub(r'' + var_to_strip + r'', '', j)
        role = re.sub(r',', '', role)
        if len(role) > 0:
            role = re.sub(r'\s+of\s*\Z', '', role)
            break
    return role


def person_set(most_recent_person,next_person):
    if len(most_recent_person)>0:
        last = most_recent_person[0]
        person = most_recent_person[1]
        role = most_recent_person[2]
        firm = most_recent_person[3]
        former = most_recent_person[4]
    else:
        last = next_person[0]
        person = next_person[1]
        role = next_person[2]
        firm = next_person[3]
        former = next_person[4]
    return last, person, role, firm, former


def former_search(sentence,role):
    former_role = 0
    former = re.findall(r'\b(?:(former\s+)\W+(?:\w+\W+){0,' + str(mxfrmer) + r'}?(' + role + r'))',sentence, re.IGNORECASE)
    if len(former)>0:
        former_role = 1
    return former_role



def make_row(last, person, role, firm, former, n, people):
    # Combine last name, person, firm, and role, for that person.
    row = [last, person, role, firm, former, n]
    # If something in each key varable then add to the dataset of people
    if min(len(row[0]), len(row[1]), len(row[2]), len(row[3])) > 0:
        # Filter out people who already were added to the list
        if last in lsts:
            pass
        else:
            people.append(row)
            # remove duplicates
            people = [list(x) for x in set(tuple(x) for x in people)]
    return people

def person_name(person):
    last = ""
    sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
    if len(sfx) > 0:
        last = person.split()[-2]
    elif len(sfx) == 0:
        last = person.split()[-1]
    return person, last

def flatten(ugly_list_of_lists):
    result = []
    for el in ugly_list_of_lists:
        if hasattr(el, "__iter__") and not isinstance(el, str):
            result.extend(flatten(el))
        else:
            result.append(el)
    return result



# function to print sentiments
# of the sentence.
def sentiment_scores(sentence):
    # Create a SentimentIntensityAnalyzer object.
    sid_obj = SentimentIntensityAnalyzer()

    # polarity_scores method of SentimentIntensityAnalyzer
    # oject gives a sentiment dictionary.
    # which contains pos, neg, neu, and compound scores.
    sentiment_dict = sid_obj.polarity_scores(sentence)
    sent_vader_neg = sentiment_dict['neg']
    sent_vader_neu = sentiment_dict['neu']
    sent_vader_pos = sentiment_dict['pos']

    return sent_vader_neg, sent_vader_neu, sent_vader_pos



z = 0


#######################################################################################
#######################################################################################
# Start of articles
#######################################################################################
#######################################################################################
#arts = range(54,85)
arts = [8]

for art in arts:
#for art in dfa.index:
    #try:
        n = 1
        h.n = 1

        t1 = time.time()
        text = str(dfa['LP'][art]) + ' ' + str(dfa['TD'][art])
        h.AN = AN = dfa["AN"][art]
        h.date = date = dfa["PD"][art]
        h.source = source = dfa["SN"][art]
        h.regions = regions = dfa["RE"][art]
        h.subjects = subjects = dfa["NS"][art]
        h.misc = misc = dfa["IPD"][art]
        h.industries = industries = dfa["IN"][art]
        h.focfirm = focfirm = dfa["Firm"][art]
        h.by = by = dfa["BY"][art]
        h.comp = comp = dfa["CO"][art]
        ##############################################################
        # Prepocessing the text to let it be analyzed by the Stanza model.
        ##############################################################

        # First, need to remove the dateline: the brief piece of text included in news articles describing
        # where and when the story was written or filed. Keeping it can erroneiously code one extra
        # relevant org such as the source (e.g., Reuters) that would mess up the attribution of quotes
        # in the code below.

        # A few ways to identify a dateline:
        d1 = re.findall(r'^.{0,' + maxch + r'}\d\d\d\d--', text, re.IGNORECASE)
        d1a = re.findall(r'^.{0,' + maxch + r'}\(.+\)\s*--', text, re.IGNORECASE)
        d1b = re.findall(r'^.{0,' + maxch + r'}\(.+\)\s*-', text, re.IGNORECASE)
        d2a = re.findall(r'^.{0,' + maxch + r'}/\b.+\b/\s*--', text, re.IGNORECASE)
        d2b = re.findall(r'^.{0,' + maxch + r'}/\b.+\b/\s*-', text, re.IGNORECASE)
        d3a = re.findall(r'^.{0,' + maxch + r'}\d\d*\s*--', text, re.IGNORECASE)
        d3b = re.findall(r'^.{0,' + maxch + r'}\d\d*\s-', text, re.IGNORECASE)
        d3c = re.findall(r'^.{0,' + maxch + r'}\d\d*,\s\w+\s*-', text, re.IGNORECASE)
        d3d = re.findall(r'^[A-Z]+\b.{0,' + maxch1 + r'}\b\s*--', text, re.IGNORECASE)
        d3e = re.findall(r'^\S+\s+-\s+', text, re.IGNORECASE)

        #print("d1", d1, "d1a", d1a,  "d1b", d1b, "d2a", d2a, "d2b", d2b,
        #      "d3a", d3a, "d3b", d3b, "d3c", d3c, "d3d", d3d)

        # Iteratively go through these instances. Only do one because could potentially
        # over cut if run on more than one.

        if len(d1) > 0:
            text = re.sub(r'^.{0,' + maxch + r'}\d\d\d\d--', '', text)

        elif len(d1a) > 0:
            text = re.sub(r'^.{0,' + maxch + r'}\(.+\)\s*--', '', text)

        elif len(d1b) > 0:
            text = re.sub(r'^.{0,' + maxch + r'}\(.+\)\s*-', '', text)

        elif len(d2a) > 0:
            text = re.sub(r'^.{0,' + maxch + r'}/\b.+\b/\s*--', '', text)

        elif len(d2b) > 0:
            text = re.sub(r'^.{0,' + maxch + r'}/\b.+\b/\s*-', '', text)

        elif len(d3a) > 0:
            text = re.sub(r'^.{0,' + maxch + r'}\d\d*\s--', '', text)

        elif len(d3b) > 0:
            text = re.sub(r'^.{0,' + maxch + r'}\d\d*\s-', '', text)

        elif len(d3c) > 0:
            text = re.sub(r'^.{0,' + maxch + r'}\d\d*,\s\w+\s*-', '', text)

        elif len(d3d) > 0:
            text = re.sub(r'^[A-Z]+\b.{0,' + maxch1 + r'}\b\s*--', '', text)

        elif len(d3e) > 0:
            text = re.sub(r'^\S+\s+-\s+', '', text)




        # Next, remove any instance of a set of non-white-space characters 1-letter parentheses (e.g., "(D)"). These mess up the NLP process as they mean something literally
        # in the PyTorch code.
        text = re.sub(r'\(\S+\)', '', text)
        # Remove anything in between angle brackets
        text = re.sub(r'<.+>', '', text)
        # Remove any instances of all caps then colon then all caps (e.g., "(NYSE: F)")
        text = re.sub(r'\(\s*[A-Z]+\s*:\s*[A-Z]+\s*\)', '', text)
        # Next, remove any of several other characters.
        text = re.sub(r'[\[\]()#/]', '', text)
        # Next replace * with period. Another issue with nlp
        text = re.sub(r'\*', r'.', text)
        # Separate any words with multiple upper-case then lower case patterns (e.g., "MatrixOne" -> "Matrix One")
        text = re.sub(r"(\b[A-Z][a-z]+)([A-Z][a-z]+)", r"\1 \2", text)

        # To correct for any errors of not spacing in articles,
        # Separate any words if said-equivalent next to other word.
        text = re.sub(r'(' + stext_split + r')([A-Z]\S+)', r"\1 \2", text)
        text = re.sub(r'(\S+)(' + stext_split + r')', r"\1 \2", text)
        # Also separate any words if any leadership role next to other word.
        text = re.sub(r'(' + leader_split + r')([A-Z]\S+)', r"\1 \2", text, re.IGNORECASE)
        text = re.sub(r'(\S+)(' + leader_split + r')', r"\1 \2", text, re.IGNORECASE)

        # Remove any words with backslashes in them. There are some articles with these for tagging purposes
        text = re.sub(r'[\\+]', '_', text)
        # Remove suffixes where a period
        text = re.sub(r'(' + suf + r')[.]', '', text)
        # Remove suffixes all together even without a period
        text = re.sub(r'(' + suf + r')', '', text)
        # Remove single capital letter followed by period
        text = re.sub(r'[A-Z][.]', '', text)
        # Replace erroneious quote then s with apostrophy s.
        text = re.sub(r'"s ', r"'s ", text)
        # Remove elipses
        text = re.sub(r'\.\s*\.\s*\.', r'', text)
        # Remove any spaces beyond one.
        text = re.sub(r'(\s+)', r' ', text)
        # Remove the title of of professor with just a the Prof
        text = re.sub(r'Prof.', r'Prof', text)



        # Next remove any quotes around a single word (e.g., "great" or "Well,"). These also irratate the nlp program
        zy = []
        t = re.findall(r'(' + start + r')(\w+\w,?)(' + end + r')', text)
        for w in t:
            u = ''.join(''.join(elems) for elems in w)
            y = re.sub(r'[' + qs + r']', '', u)
            z = [u, y]
            zy.append(z)
        for w in zy:
            text = re.sub(r'(' + w[0] + r')', w[1], text)


        # Next replace multi-dashed words with no dashes.
        zy = []
        t = re.findall(r'\w+-\w+-\w+', text)
        for w in t:
            u = ''.join(''.join(elems) for elems in w)
            y = re.sub(r'-', '', u)
            z = [u, y]
            zy.append(z)
        for w in zy:
            text = re.sub(r'(' + w[0] + r')', w[1], text)

        # Remove any commas before suffixes.
        zy = []
        t = re.findall(r',\s*' + suf + r'[^a-z]', text)
        for w in t:
            u = ''.join(''.join(elems) for elems in w)
            y = re.sub(r',', '', u)
            z = [u, y]
            zy.append(z)
        for w in zy:
            text = re.sub(r'(' + w[0] + r')', w[1], text)



        #Encode the text as English to remove any foreign letters
        encoded_string = text.encode("ascii", "ignore")
        text = encoded_string.decode()

        ################################
        # Loading the text into the Stanza nlp system
        ################################
        doc = nlp(text)

        #try:
        people = []
        sppl = []

        fulls = []
        lsts = []

        ppl_inrow = []

        #Set the sentence count to 0
        s = 0
        #Create a count variable for each article that starts at 0 and +1 every time you add to the DF
        dx = 0

        corgs = []

        # Create a sentence-level DFs
        columns = ['no.', 'sentence', 'filtered_sent', 'sentiment_stanza', 'orgs', 'corgs', 'ppl', 'prows', 'people',
                   'sppl', 'covered', 'events', 'sent_vader_neg', 'sent_vader_neu', 'sent_vader_pos']
        dfsent = pd.DataFrame(columns=columns)

        # Create a DF for people
        columns = ['first_sent', 'last', 'full', 'role', 'firm', 'former']
        dfpeople = pd.DataFrame(columns=columns)

        # Look through document and see if has instances of at least four all-cap words
        # broken up by non-word characters (e.g., spaces, commas) followed by a colon.
        f = re.findall(r'(((\b[A-Z]+\b\W*){4,}):)', text)
        comma_string = 0
        if len(f)> 0:
            comma_string = re.findall(r',', f[0][0])
        # If there is at least 1 of these very unique instances (with commas), this is got to
        # be a transcript of some kind.
        if len(f) > 0 and len(comma_string) > 0:

            # Go through each sentence to find if it's introducing a person and their role and firm
            for sent in doc.sentences:

                stence = sent.text
                # Make sure there is something in the sentence. If not, skip.
                op = re.findall(r'^OPERATOR', stence)
                if len(op) == 0:

                    # Set a trigger variable to 0 once start a new sentence.
                    dr = 0
                    # Calculate the sentiment for each sentence
                    h.sentiment_stanza = sent.sentiment
                    h.sent_vader_neg = sentiment_scores(stence)[0]
                    h.sent_vader_neu = sentiment_scores(stence)[1]
                    h.sent_vader_pos = sentiment_scores(stence)[2]

                    #for ent in sent.ents:
                    #    print(n, ent.text, ent.type)

                    last = ""
                    role = ""
                    role1 = ""
                    person = ""
                    firm = ""
                    firm1 = ""
                    # In these conference calls, only current roles speak.
                    # Plus, it wouldn't say "FORMER CEO..." anyway
                    former = 0
                    # Check if name title in the sentence:
                    g = re.findall(r'(((\b[A-Z]+\b\W*){4,}):)', stence)
                    if len(g) > 0:
                        a = g[0][0]
                        a = re.sub(r':', '', a)
                        # If company name is Inc and has a comma, remove the comma
                        a = re.sub(',\s+(INC|LLC)', '', a, re.IGNORECASE)
                        p = a.split(',')

                        person = p[0].title()
                        last = person_name(person)[1].title()
                        firm1 = p[-1]
                        # Lastly, if there are four pieces the role is the 2nd and 3rd pieces so combine
                        if len(p) ==5:
                            r = p[1] + p[2] + p[3]
                            role1 = ''.join(''.join(elems) for elems in r)
                        elif len(p) == 4:
                            r = p[1] + p[2]
                            role1 = ''.join(''.join(elems) for elems in r)
                        elif len(p) == 3:
                            role1 = p[1]
                        role_check = rolesearch(role1)
                        if len(role_check)==0:
                            role = firm1
                            firm = role1.title()

                        elif len(role_check)>0:
                            role = role1
                            firm = firm1.title()


                        row = [last, person, role, firm, former, n]

                        # If something in each key varable then add to the dataset of people
                        if min(len(row[0]), len(row[1]), len(row[2]), len(row[3])) > 0:
                            # Filter out people who already were added to the list
                            if last in lsts:
                                pass
                            else:
                                people.append(row)
                                # remove duplicates
                                people = [list(x) for x in set(tuple(x) for x in people)]
                                sppl.append(row)
                                lsts.append(last)



                    elif len(g) == 0:
                        m = []
                        # Check if previously-introduced person is talking
                        for per in people:
                            last = per[0]
                            last_string = re.findall(r'' + last + r':', stence, re.IGNORECASE)
                            m = m + last_string
                        m = ''.join(''.join(elems) for elems in m)
                        if len(m) > 0:
                            for per in people:
                                last = per[0]
                                l = re.findall(r'' + last + r':', stence, re.IGNORECASE)
                                if len(l) > 0:
                                    # Strip the beginning with the quote.
                                    h.quote = re.sub(r'^.*' + last + r':', '', stence)
                                    h.sentence = h.quote
                                    h.last = per[0]
                                    h.person = per[1]
                                    h.role = per[2]
                                    h.firm = per[3]
                                    h.former = per[4]
                                    h.qtype = "Transcript"
                                    h.qsaid = "yes"
                                    h.qpref = "yes"
                                    h.comment = "Transcript"
                                    h.todf()
                                    dx += 1
                                    # Add this person to the running list of sequential people speaking
                                    sppl.append(per)
                                    break


                        # If there isn't an introduction of a person or a reintroduction of a person
                        elif len(m) == 0:
                            # Are there any people in the people list yet?
                            # If so, then the sentence belongs to the last person assigned.
                            if len(people) > 0:
                                h.quote = re.sub(r'^.*' + last + r':', '', stence)
                                h.sentence = h.quote
                                h.last = sppl[-1][0]
                                h.person = sppl[-1][1]
                                h.role = sppl[-1][2]
                                h.firm = sppl[-1][3]
                                h.former = sppl[-1][4]
                                h.qtype = "Transcript"
                                h.qsaid = "yes"
                                h.qpref = "yes"
                                h.comment = "Transcript"
                                h.todf()
                                dx += 1
                n += 1
                h.n += 1

        #####################
        # Non-transcripts (vast majority of instances)
        #####################
        else:
            n =  1
            h.n = 1
            ########################################################
            # Build the person database of entire article. You will do this before iteritively going through
            # each sentence to check for any person through last name, title or pronoun.
            ########################################################
            for sent in doc.sentences:

                # reset key variables for each sentence:
                quote = ""
                last = ""
                person = ""
                firm = ""
                role = ""
                former = ""
                qtype = ""
                qsaid = ""
                qpref = ""
                comment = ""

                ppl = []
                lasts = []
                prod = []
                orgs = []
                events = []

                prows = []

                #variables for future sentences as needed (for multi-sentence quotes)
                fppl = []
                flasts = []
                forgs = []
                fprows = []
                fprod = []

                h.sentence = stence = sent.text

                # Calculate the sentiment for each sentence
                h.sentiment_stanza = sentiment_stanza = sent.sentiment
                h.sent_vader_neg = sent_vader_neg = sentiment_scores(stence)[0]
                h.sent_vader_neu = sent_vader_neu = sentiment_scores(stence)[1]
                h.sent_vader_pos = sent_vader_pos = sentiment_scores(stence)[2]

                #for ent in sent.ents:
                #    print(n, ent.text, ent.type)

                ############################################
                # CREATE STRING THAT FILTERS OUT ANY QUOTES:
                ############################################
                non_quotes = []
                # split the sentence by quotation mark
                quote_split = re.split(r'(' + qs + r')', stence)
                # Remove blanks if any
                quote_split = list(filter(None, quote_split))
                # Remove any items that are quotation marks
                rex = re.compile(r'' + qs + r'')
                quote_split = [x for x in quote_split if not rex.match(x)]

                if len(quote_split)==1:
                    filtered_sent = stence

                else:
                    for q in range(len(quote_split)):
                        quote_beginning = re.findall(r'^\s*(' + start + r')', stence, re.IGNORECASE)
                        if len(quote_beginning) > 0:
                            # If quote from the start of sentence,
                            # Want position 0, 2, etc. whether it's even or odd number of quotation marks
                            if q % 2 == 1:
                                non_quotes.append(quote_split[q])
                        elif len(quote_beginning) == 0:
                            # Now need to distinquish between whether the first quote is the end of a multi-sentence
                            # or the beginning of a new one. A space before the first quote should distinguish this.
                            before_first_quote = re.findall(r'^.+?(?=' + start + r')', stence, re.IGNORECASE)
                            if len(before_first_quote) > 0:
                                space_before_first_quote = re.findall(r'\s$', before_first_quote[0], re.IGNORECASE)
                                if len(space_before_first_quote) > 0:
                                    # Check if a space before the first quotation mark. If so, it would indicate it's not
                                    # The end of a multi-sentence quote.
                                    if q % 2 == 0:
                                        non_quotes.append(quote_split[q])
                                else:
                                    if q % 2 == 1:
                                        non_quotes.append(quote_split[q])

                    filtered_sent = ''.join(''.join(elems) for elems in non_quotes)


                ############################################
                # Check if you can add to dataset of people/firm/role
                # First, compile a list of people and organizations and products iteratively within the sentence.
                # You use products list in the orgs list so need to run them sep.
                ############################################
                for ent in sent.ents:
                    if ent.type == "PERSON":
                        # Check if inside the filtered portion of sentence
                        z = re.findall(r'' + ent.text + r'', filtered_sent)
                        if len(z) > 0:
                            #Make sure it's not a car company that looks like a name e.g., Aston Martin
                            car_maker_check = re.findall(r'' + car_makers_names + r'', ent.text)
                            if len(car_maker_check)>0:
                                orgs.append(ent.text)
                                corgs.append(ent.text)
                            else:
                                ppl.append(ent.text)
                                sfx = re.findall(r'(' + suf + r')', ent.text, re.IGNORECASE)
                                if len(sfx) > 0:
                                    l = ent.text.split()[-2]
                                    lasts.append(l)
                                elif len(sfx) == 0:
                                    l = ent.text.split()[-1]
                                    lasts.append(l)

                    # Check if a last name was incorrectly marked as not a PERSON (e.g., an ORG)
                    # and if so, assign as PERSON
                    if len(people) > 0:
                        for per in people:
                            last = per[0]
                            last_said = re.findall('(' + last + r')', filtered_sent, re.IGNORECASE)
                            last_in_sent = ''.join(''.join(elems) for elems in last_said)
                            if len(last_in_sent) > 0:
                                ppl.append(per[1])

                # Next the products, to be used for the orgs
                for ent in sent.ents:
                    if ent.type == "PRODUCT":
                        pduct = ent.text
                        # remove odd punctuation such as ")" that can mess up the regex down the line
                        pduct = re.sub(r'[()]', '', pduct)
                        z = re.findall(r'' + pduct + r'', filtered_sent)
                        if len(z) > 0:
                            prod.append(pduct)
                # Finally, the orgs
                for ent in sent.ents:
                    if ent.type == "ORG":
                        tion = ent.text
                        # remove odd punctuation such as ")" that can mess up the regex down the line
                        tion = re.sub(r'[()]', '', tion)
                        # Check if in the media, if so don't include
                        news_check = re.findall(r'(' + media + r')', tion)
                        if len(news_check)==0:
                            z = re.findall(r'' + tion + r'', filtered_sent)
                            if len(z) > 0:
                                # don't include acronyms
                                a = re.findall(r'\(\s*' + tion + r'\s*\)', stence)
                                a = ''.join(''.join(elems) for elems in a)
                                if len(a) == 0:
                                    # Not an acronym, so see if a car type is followed by the org
                                    # If the org name is followed closely by a vehicle type name (e.g., sedan)
                                    # The name should not be considered in terms of people and roles.
                                    # It would just be a make of a car, e.g., "Mercedes E-class sedan"
                                    car_make = re.findall(
                                        r'(' + tion + r')\W+(?:\w+\W+){0,' + str(mxcar) + r'}?(' + car_types + r')',
                                        filtered_sent, re.IGNORECASE)
                                    if len(car_make) == 0:
                                        # Not a car make, so see if any products in sentence:
                                        # If not, see if already in list
                                        if len(prod) == 0:
                                            if tion not in orgs:
                                                # If not, append
                                                orgs.append(tion)
                                                corgs.append(tion)
                                        # If products in the sentence,
                                        # check if it's not just a brand name with a product after it
                                        # like "Toyota SRV"
                                        else:
                                            d = []
                                            for p in prod:
                                                c = re.findall(r'' + ent.type + r'\s+' + p + r'', stence)
                                                d = d + c
                                            # If not, then just make sure it's not already in the list.
                                            if len(d) == 0:
                                                if tion not in orgs:
                                                    # If not, append
                                                    orgs.append(tion)
                                                    corgs.append(tion)

                # also, add the events
                for ent in sent.ents:
                    if ent.type == "EVENT":
                        event = ent.text
                        events.append(event)

                # remove duplicates from each list:
                orgs = list(set(orgs))
                corgs = list(set(corgs))
                lasts = list(set(lasts))
                prod = list(set(prod))
                ppl = list(set(ppl))


                # Next, if there are any people, see if should add to running People list
                # First, check if at least one person
                if len(ppl) > 0:

                    for per in ppl:
                        # see if the person's title is mentioned
                        # reset key variables each person reference

                        last = ""
                        person = ""
                        firm = ""
                        role = ""
                        former = ""
                        if len(ppl) == 1:
                            # If exactly 1 org, not matter where the person is in proximity of title/role
                            if len(orgs) == 1:
                                role_count = rolesearch(filtered_sent)
                                if len(role_count)>1:
                                    role_list = shortrolesearch(filtered_sent, per, mxld)
                                else:
                                    role_list = longrolesearch(filtered_sent, per)
                            # If either 0 or >1 orgs, check if person and title are close
                            else:
                                role_list = shortrolesearch(filtered_sent, per, mxld)


                        # Then by definition, must be >1 people
                        else:
                            role_list = shortrolesearch(filtered_sent, per, mxlds)


                        # If the title is mentioned, move to next step
                        if len(role_list) > 0:

                            # If there is just one organization mentioned, assign
                            if len(orgs) == 1:

                                org = orgs[0]
                                firm = org.replace(r"'s", "")
                                person = per
                                last = person_name(person)[1]
                                role = role_set(role_list, per)
                                former = former_search(filtered_sent, role)

                            # Multiple orgs in sentence
                            elif len(orgs) > 1:



                                # Create list and variable that will be used at the end of this iteration of orgs
                                # "gm" checks if there are any instances in which an org is a substring of another
                                #  org name. If yes, len(g)>0. If not, len(g) = 0 and you're good.
                                c = 0
                                b = 0
                                q = 0
                                xl = 0
                                op = ""
                                gh = []
                                gp = []
                                rolexl = ""
                                firmxl = ""
                                for o in orgs:

                                    or1 = orgs.copy()
                                    or1.remove(o)

                                    li = [match for match in or1 if o in match]
                                    gh.append(li)
                                    # If there is substring occurance, create a list with all firms
                                    # with that substring

                                    if len(li) > 0:
                                        li.append(o)
                                        # If gp not a list, then turn it into a list
                                        gp = [gp] if isinstance(gp, str) else gp
                                        gp.append(li)
                                        gp = gp[0]

                                # In the case of multiple substrings of orgs (e.g., Honda and Honda US,
                                # and Toyota and Toyota Europe), the gh list of lists can get really ugly and cause
                                # problems. So you need to 'flatten' the list so each element is at level 1.
                                gh = flatten(gh)

                                gm = ''.join(''.join(elems) for elems in gh)

                                repeat_orgs = []
                                repeat_count = 0
                                for o in orgs:
                                    # We want to make sure that the right org name gets assigned. One problem is
                                    # if there is a org that is a substring of the correct org name such as GM within the
                                    # correct GM Europe in (e.g., "GM is working on developing the new model, said
                                    # Carl-Peter Forster, president of GM Europe"). In these cases, it would assign "GM"
                                    # and not GM Europe. So, we need to set a couple of variables; namely, c if we've assigned
                                    # already, and b if the focal org is a substring of another org in the sentence.
                                    # So, if you've already assigned (c>0) and a subset of another org (b>0), then
                                    # you're done. Four cases: (1) GM before GM Europe and GM Europe is correct:
                                    # step 1, c = 0 and b>1 GM is focal org and assigns first, then when do GM Europe
                                    # c = 1 and b = 0 so that ultimately gets assigned. (2) GM before GM Europe and GM
                                    # is correct: GM gets assigned as above and when GM Europe goes through it doesn't
                                    # replace GM. (3) GM Europe before GM and GM Europe correct: GM Europe assigned
                                    # then under GM c = 1 and b = 1 so stay at GM Europe.. (4) GM Europe before GM
                                    # and GM is correct: GM Europe doesn't match, then under GM, c = 0 and b = 1
                                    # so GM is assigned.

                                    # To do so, we need to make a list with the focal org omitted
                                    or2 = orgs.copy()
                                    or2.remove(o)
                                    b = [match for match in or2 if o in match]
                                    if len(b)>0:
                                        repeat_orgs.append(b)

                                for o in orgs:

                                    #rerun this same process because now have the total b list.
                                    or2 = orgs.copy()
                                    or2.remove(o)
                                    b = [match for match in or2 if o in match]

                                    if len (b)>0:
                                        repeat_count +=1


                                    if c > 0 and len(b) > 0:
                                        break
                                    # Consider that there might be 3+ firms in the sentence, with one with a substring
                                    # of another. If it has gone through the two related versions and assigned, this
                                    # code will prevent it from looking at the third+ firm which it will not assign.
                                    elif c == repeat_count + 1 and len(repeat_orgs) > 0 and repeat_count == len(repeat_orgs):
                                        break

                                    else:
                                        # Check if focal org is in the list of multi versions (e.g., GM, GM Europe)
                                        # variable "l" set to 0, and turns 1 if in the list.
                                        l = 0
                                        if o in gp:
                                            l = 1

                                        # There are strong matches (e.g., whose, 's, and of) in which if we find a
                                        # match the only thing that would make us look further is if it's in the multi-
                                        # version list (e.g., GM, GM Europe). If not, you're done. The h variable
                                        # will > 0 once a strong match has occurred. If h > 0 (strong match has
                                        # occurred) but l == 0 (the focal org is not in the multi-version list)
                                        # skip ("continue") this org to the next org in the for loop.
                                        # Essentially, at this point you'll only update a strong match
                                        # with a better version of the org for which you had a strong match.

                                        if q > 0 and l == 0:
                                            continue

                                        # Check if there is possession with one "'s" in the org
                                        o1 = re.findall(r'' + o + r'\'s', filtered_sent, re.IGNORECASE)
                                        o2 = re.findall(r'\'s', o, re.IGNORECASE)

                                        og = o1 + o2

                                        if len(og) > 0:
                                            if len(ppl)==1:
                                                os = longrolesearch(filtered_sent, o)
                                            else:
                                                os = shortrolesearch(filtered_sent, o, mxorg1)
                                            om = ''.join(''.join(elems) for elems in os)

                                            if len(os) > 0:
                                                person = per
                                                last = person_name(person)[1]
                                                role = role_set(os, o)
                                                o = o.replace(r"'s", "")
                                                firm = o
                                                former = former_search(filtered_sent, role)

                                                c += 1
                                                q += 1
                                                if len(gm) == 0:
                                                    break

                                        # If not, check if there is possession with "whose"
                                        # Org then whose then title.
                                        elif len(og) == 0:
                                            m = whoserolesearch(filtered_sent, o, mxorg1)

                                            role = ""
                                            # If whose then title, assign
                                            if len(m) > 0:
                                                person = per
                                                last = person_name(person)[1]
                                                role = role_set(m, o)
                                                firm = o
                                                former = former_search(filtered_sent, role)
                                                c += 1
                                                q += 1
                                                if len(gm) == 0:
                                                    break






                                            # If not, check if [title] "of" near [org]
                                            elif len(m) == 0:
                                                m = ofrolesearch(filtered_sent, o, mxorg1)

                                                j = ''.join(''.join(elems) for elems in m)

                                                #If so, assign
                                                if len(j) > 0:
                                                    person = per
                                                    last = person_name(person)[1]
                                                    role = m[0][-1]
                                                    firm = o
                                                    former = former_search(filtered_sent, role)
                                                    c += 1
                                                    q += 1
                                                    if len(gm) == 0:
                                                        break

                                                # If not, check if "company's" then title. Then just take the
                                                # first company because both companies would be related.
                                                elif len(j) == 0:
                                                    m = companysrolesearch(filtered_sent, o, mxcomp)
                                                    k = ''.join(''.join(elems) for elems in m)
                                                    if len(k)>0 and len(ppl) == 1:
                                                        person = per
                                                        last = person_name(person)[1]
                                                        role = m[0][-1]
                                                        firm = o
                                                        former = former_search(filtered_sent, role)
                                                        c += 1
                                                        q += 1
                                                        if len(gm) == 0:
                                                            break

                                                    # If not, check if an org has a word indicated the object and not the
                                                    # subject of the sentence. For example, consider the word "with" before VW in
                                                    # "Ford entered a JV with VW, said John Smith VP of Sales." In this case
                                                    # VW is irrelevant in terms of assigning the company to the person.
                                                    elif len(k) == 0:
                                                        if len (orgs) == 2:
                                                            orgs2 = orgs.copy()
                                                            for i, org in enumerate(orgs2):
                                                                check_with = re.findall(r'(' + prepositions + r')\s+' + org + r'',
                                                                                        filtered_sent, re.IGNORECASE)
                                                                if len(check_with) > 0:
                                                                    del orgs2[i]
                                                            if len(orgs2)==1:
                                                                os = longrolesearch(filtered_sent, orgs2[0])
                                                                person = per
                                                                last = person_name(person)[1]
                                                                firm = orgs2[0]
                                                                role = role_set(os, firm)
                                                                former = former_search(filtered_sent, role)

                                                                c += 1
                                                                q += 1
                                                                if len(gm) == 0:
                                                                    break



                                                            # If only one person and the only multiple companies are
                                                            # part of each other, just pick the first one mentioned.
                                                            elif len(orgs2)!= 1:
                                                                if len(ppl) == 1 and repeat_count+1 == len(orgs):
                                                                    os = longrolesearch(filtered_sent, o)
                                                                    person = per
                                                                    last = person_name(person)[1]
                                                                    role = role_set(os, o)
                                                                    firm = o
                                                                    former = former_search(filtered_sent, role)

                                                                    c += 1
                                                                    q += 1
                                                                    if len(gm) == 0:
                                                                        break

                                                                # If not any of these three strong matches ('s, of, whose),
                                                                # check on a weaker match: if name are close to title
                                                                else:

                                                                    m = shortrolesearch(filtered_sent, o, mxorg1)
                                                                    on = ''.join(''.join(elems) for elems in m)

                                                                    # If so, check if person is also near the title
                                                                    if len(on) > 0:
                                                                        m = shortrolesearch(filtered_sent, per, mxorg1)
                                                                        ol = ''.join(''.join(elems) for elems in m)

                                                                        # If name and person are close to title, assign
                                                                        if len(ol) > 0:
                                                                            person = per
                                                                            last = person_name(person)[1]
                                                                            role = role_set(m, per)
                                                                            former = former_search(filtered_sent, role)
                                                                            firm = o
                                                                            c += 1
                                                                            if len(gm) == 0:
                                                                                break
                                                                    # If not, check if name followed by just one org and a title.
                                                                    elif len(on) == 0:
                                                                        m = threerolesearch(filtered_sent, per, o)

                                                                        if len(m) > 0:
                                                                            r = role_set(role_list,per)
                                                                            if len(r) > 0:
                                                                                xl += 1
                                                                                rolexl = r
                                                                                firmxl = o
                                                                                c += 1


                                if xl == 1:
                                    role = rolexl
                                    former = former_search(filtered_sent, role)
                                    firm = firmxl
                                    person = per
                                    last = person_name(person)[1]


                            elif len(orgs) == 0:

                                # Check if an analyst from an unnamed firm
                                analyst_search = re.findall(r'(' + analyst + r')', filtered_sent, re.IGNORECASE)
                                if len(analyst_search)>0:
                                    person = per
                                    last = person_name(person)[1]
                                    role = role_set(role_list, per)
                                    former = former_search(filtered_sent, role)
                                    firm = "NOTNAMED"

                                else:
                                    #Check if there was an event name in the previous sentence.
                                    if n>1 and len(dfsent.iloc[n-2]['events'])>0:
                                        person = per
                                        last = person_name(person)[1]
                                        role = role_set(role_list, per)
                                        former = former_search(filtered_sent, role)
                                        firms = dfsent.iloc[n-2]['events']

                                # If there's a titled person with no org, it may be the org in
                                # previous reference. If a repeat person, this will be filtered out below.
                                # If not org in previous sentence, this will also be dropped:
                                if len(dfsent) > 0:
                                    person = per
                                    last = person_name(person)[1]
                                    role = role_set(role_list, per)
                                    former = former_search(filtered_sent, role)
                                    firms = dfsent.iloc[n-2]['orgs']
                                    if len(firms) == 1:
                                        firm = firms[0]
                        # If there's no leadership position, check if there is a person
                        # If there is and it says "of" something, add it to the people list
                        elif len(role_list) == 0:
                            if len(ppl) == 1:
                                if len(orgs) > 0:
                                    for org in orgs:
                                        q = re.findall(r'.+' + ppl[0] + r'\s+(of|at)\s+' + org + r'', filtered_sent,
                                                       re.IGNORECASE)
                                        if len(q) > 0:
                                            firm = org
                                            role = "NA"
                                            former = "NA"
                                            person = per
                                            last = person_name(person)[1]

                        people = make_row(last, person, role, firm, former, n, people)

                elif len(ppl) == 0:
                    if len(orgs) == 1:
                        o = orgs[0]
                        m = longrolesearch(filtered_sent, o)

                        on = ''.join(''.join(elems) for elems in m)

                        # If there is a role and org close together, then check if a single person
                        # in the next sentence:
                        if len(on) > 0 and n<len(doc.sentences):
                            pz = []
                            j = doc.sentences[n]
                            b = j.text
                            #### Filter out any quotes, then use that as the reference for filtered_sent below
                            # Must stip the the sentence of any quotes in order to arrive at appropriate orgs, people, and products
                            # First, remove any information within quotes:
                            sty1 = re.sub(r'(' + start + r').+(' + end + r')', ' ', b)
                            # Second, remove information in beginning of multi-sentence quotes
                            # the start quote would have a space before the quote
                            sty2 = re.sub(r'([:|,])\s+(' + start + r')\S.+$', ' ', sty1)
                            # Third, remove cases in which the quote begins in the sentance and doesn't end.
                            sty3 = re.sub(r'^(' + start + r')\S.+$', ' ', sty2)
                            # Last, remove information at the end of multi-sentence quotes
                            # the end quote would have a space after the quote
                            sty4 = re.sub(r'^.+(' + end + r')\s+', ' ', sty3)

                            # Check if there are any people in the non-quote portion of the sentence:
                            for ent in j.ents:
                                if ent.type == "PERSON":
                                    z = re.findall(r'' + ent.text + r'', sty4)
                                    if len(z) > 0:
                                        # This will add the person and last name to these lists
                                        pz.append(ent.text)

                            # If there's just one person, assign
                            if len(pz) == 1:
                                # Add to the ppl and lasts lists
                                person = pz[0]
                                sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                                if len(sfx) > 0:
                                    last = person.split()[-2]
                                    lasts.append(last)

                                elif len(sfx) == 0:
                                    last = person.split()[-1]
                                    lasts.append(last)

                                firm = o
                                # Add the role by iteratively going through the hierarchy of titles
                                role = role_set(m, firm)
                                former = former_search(filtered_sent, role)
                                people = make_row(last, person, role, firm, former, n, people)
                                # In order to include this person in the sppl list, need to add it
                                # right here. It won't be added below
                                # because the person's last name is not in the focal sentence,
                                # which is the condition to be added for all other instances.
                                sppl.append(people[-1])

                    elif len(orgs)>1:
                        # If there is a role and org close together, then check if a single person
                        # in the next sentence:
                        for o in orgs:
                            m = shortrolesearch(filtered_sent, o, mxorg1)

                            on = ''.join(''.join(elems) for elems in m)
                            if len(on) > 0 and n < len(doc.sentences):
                                pz = []
                                j = doc.sentences[n]
                                b = j.text
                                #### Filter out any quotes, then use that as the reference for filtered_sent below
                                # Must stip the the sentence of any quotes in order to arrive at appropriate orgs, people, and products
                                # First, remove any information within quotes:
                                sty1 = re.sub(r'(' + start + r').+(' + end + r')', ' ', b)
                                # Second, remove information in beginning of multi-sentence quotes
                                # the start quote would have a space before the quote
                                sty2 = re.sub(r'([:|,])\s+(' + start + r')\S.+$', ' ', sty1)
                                # Third, remove cases in which the quote begins in the sentance and doesn't end.
                                sty3 = re.sub(r'^(' + start + r')\S.+$', ' ', sty2)
                                # Last, remove information at the end of multi-sentence quotes
                                # the end quote would have a space after the quote
                                sty4 = re.sub(r'^.+(' + end + r')\s+', ' ', sty3)

                                # Check if there are any people in the non-quote portion of the sentence:
                                for ent in j.ents:
                                    if ent.type == "PERSON":
                                        z = re.findall(r'' + ent.text + r'', sty4)
                                        if len(z) > 0:
                                            # This will add the person and last name to these lists
                                            pz.append(ent.text)

                                # If there's just one person, assign
                                if len(pz) == 1:
                                    # Add to the ppl and lasts lists
                                    person = pz[0]
                                    sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                                    if len(sfx) > 0:
                                        last = person.split()[-2]
                                        lasts.append(last)

                                    elif len(sfx) == 0:
                                        last = person.split()[-1]
                                        lasts.append(last)

                                    firm = o
                                    # Add the role by iteratively going through the hierarchy of titles
                                    role = role_set(m, firm)
                                    former = former_search(filtered_sent, role)
                                    people = make_row(last, person, role, firm, former, n, people)
                                    # In order to include this person in the sppl list, need to add it
                                    # right here. It won't be added below
                                    # because the person's last name is not in the focal sentence,
                                    # which is the condition to be added for all other instances.
                                    sppl.append(people[-1])

                    elif len(orgs)==0:
                        # For the situation in which a role is named but no name or affiliation such as when
                        # an "analyst" does not want to be named, we still need to capture them
                        role = role_set(rolesearch(filtered_sent), '')
                        former = former_search(filtered_sent, role)
                        k1 = re.findall(
                            r'(' + role + r')\W+(?:\w+\W+){0,' + str(mxpron) + r'}?(' + unnamed + r')',
                            filtered_sent, re.IGNORECASE)
                        k2 = re.findall(
                            r'(' + unnamed + r')\W+(?:\w+\W+){0,' + str(mxpron) + r'}?(' + role + r')',
                            filtered_sent, re.IGNORECASE)
                        is_misc_role = k1 + k2
                        if len(is_misc_role) > 0:
                            firm = "NOTNAMED"
                            person = "NOTNAMED"
                            last = role
                            people = make_row(last, person, role, firm, former, n, people)

                        elif len(is_misc_role) == 0:
                            # if it's past the first row, Check if people listing that hasn't been
                            # placed in people yet
                            if n>1:
                                for sindex in range(0, n):
                                    # manipulate the count to start on this row and move backwards
                                    # (up to top of the df)
                                    unassigned_last = ""
                                    unassigned_person = []
                                    index = int(n-2-sindex)
                                    unassigned_persons = dfsent.iloc[index]['ppl']
                                    people_list = dfsent.iloc[index]['people']
                                    if len(unassigned_persons)>0:
                                        unassigned_person = unassigned_persons[0]
                                        sfx = re.findall(r'(' + suf + r')',  unassigned_person, re.IGNORECASE)
                                        if len(sfx) > 0:
                                            unassigned_last = unassigned_person.split()[-2]
                                        elif len(sfx) == 0:
                                            unassigned_last = unassigned_person.split()[-1]

                                        if len(unassigned_person) == 0:
                                            continue
                                        else:
                                            if unassigned_last in lsts:
                                                continue
                                            else:
                                                role = role_set(rolesearch(filtered_sent), '')
                                                former = former_search(filtered_sent, role)
                                                if len(role)>0:
                                                    firm = "NOTNAMED"
                                                    person = unassigned_person
                                                    last = unassigned_last
                                                    people = make_row(last, person, role, firm, former, n, people)
                                                    break

                # Running list of full names of People list of lists
                fulls = [per[1] for per in people]
                lsts = [per[0] for per in people]

                # Create a list of people referenced in the focal sentence.
                if len(people) > 0:
                    for per in people:
                        last = per[0]
                        m = re.findall(r'\b' + last + r'\b', filtered_sent, re.IGNORECASE)
                        if len(m) > 0:
                            prow = [per[0], per[1], per[2], per[3], per[4], n]
                            prows.append(prow)
                            dfpeople.loc[len(dfpeople.index)] = [n, per[0], per[1], per[2], per[3], per[4]]
                            dfpeople = dfpeople.drop_duplicates(subset=['last', 'full'], keep='first')\
                                .reset_index(drop=True)


                # Cumulative list of people from people list referenced in sentences
                # Will refer to this in cases of pronouns (e.g., "she said"), etc.
                for per in people:
                    y = re.findall(r'\b' + per[0] + r'\b', filtered_sent, re.IGNORECASE)
                    if len(y) > 0:
                        per[5] = n
                        sppl.append(per)

                # Carry forward the prows column until replaced with next person
                # In essence, prows tells you the most recent person in each sentence including the focal one.
                if len(prows)==0 and len(dfsent.index)>0:
                    prows = dfsent.loc[len(dfsent.index) - 1].at['prows']

                covered = None

                dfsent.loc[len(dfsent.index)] = [n, stence, filtered_sent, sentiment_stanza, orgs, corgs, ppl,
                                                 prows, people, sppl, covered, events, sent_vader_neg, sent_vader_neu,
                                                 sent_vader_pos]

                n += 1
                h.n += 1

            ############################################################################################
            ############################################################################################
            # NOW YOU HAVE THE LIST, TAG SENTENCES WITH PEOPLE AND QUOTES
            ############################################################################################
            ############################################################################################
            # Reset count variables
            n = 1
            h.n = 1
            for se in dfsent.index:

                # Assign the various lists and variables for this sentence from the dfsent dataframe
                stence = dfsent["sentence"][se]
                h.sentence = stence
                filtered_sent = dfsent["filtered_sent"][se]
                h.sentiment_stanza = dfsent["sentiment_stanza"][se]
                h.sent_vader_neg = dfsent["sent_vader_neg"][se]
                h.sent_vader_neu = dfsent["sent_vader_neu"][se]
                h.sent_vader_pos = dfsent["sent_vader_pos"][se]
                orgs = dfsent["orgs"][se]
                corgs = dfsent["orgs"][se]
                ppl = dfsent["ppl"][se]
                prows = dfsent["prows"][se]
                people = dfsent["people"][se]
                events = dfsent["events"][se]



                if dfsent["covered"][se] == 1:
                    info_extraction(stence, orgs)
                    n += 1
                    h.n += 1
                    continue


                # Deploy function to extract information from the sentence
                info_extraction(stence, orgs)

                most_recent_person = []
                next_person = []
                # Sent variable to zero that there are at least one prows in dsent
                # update below if not.
                no_prows = 0
                no_people = 0
                # Establish the last person referenced (first filter later ones and then take the last one)
                if len(prows)>0:
                    most_recent_person = prows[0]



                # Establish next person referenced in case there isn't anyone in prows yet.
                # This is important for statements at the beginning of the article before designating
                # who the person is.
                elif len(prows)==0:
                    for sindex in range(n, len(doc.sentences)):
                        next_prows = dfsent['prows'][sindex]
                        if len(next_prows) > 0:
                            next_person = next_prows[0]
                            break
                    # If no prows in entire dsent, then create variable
                    df_prows = dfsent[dfsent["prows"].str.len() != 0]
                    if len(df_prows.index)==0:
                        no_prows = 1



                df_people = dfsent[dfsent["people"].str.len() != 0]
                if len(df_people.index) == 0:
                    no_people = 1


                #Set variables for sentence to use later as needed
                last_sindex = quote_set(stence)[0]
                msent_quote = quote_set(stence)[1]
                last_after_quote = quote_set(stence)[2]
                first_said = quote_set(stence)[3]
                first_pron = quote_set(stence)[4]
                first_ppl = quote_set(stence)[5]
                last_said = quote_set(stence)[6]
                last_pron = quote_set(stence)[7]
                last_ppl = quote_set(stence)[8]
                msent_quotes = quote_set(stence)[9]
                stences = quote_set(stence)[10]
                sentiments = quote_set(stence)[11]
                nindex = quote_set(stence)[12]
                quotes = quote_set(stence)[13]
                multi_sent_indicator = quote_set(stence)[14]
                sent_vader_negs = quote_set(stence)[15]
                sent_vader_neus = quote_set(stence)[16]
                sent_vader_poss = quote_set(stence)[17]



                psent_index = n - 2


                ##################################
                # SENTENCES WITH A "SAID" EQUIVALENT
                ##################################
                s = re.findall(r'' + stext + '', filtered_sent, re.IGNORECASE)

                # First check if "said" equivalent in sentence
                if len(s) > 0:
                    ##################################
                    # NO PEOPLE REFERENCE
                    ##################################
                    if len(ppl) == 0:
                        x1 = re.findall(r'(' + start + ')(.+)(' + end + ')', stence, re.IGNORECASE)
                        x1 = ''.join(''.join(elems) for elems in x1)
                        x1 = re.sub(r'(' + qs + r')', '', x1)

                        if len(x1) > 0:
                            # if there is a said equivalent and quote but no person, check if
                            # pronoun near said equivalent
                            k1 = re.findall(
                                r'(' + stext + r')\W+(?:\w+\W+){0,' + str(mxpron) + r'}?(' + pron + r')',
                                filtered_sent, re.IGNORECASE)
                            k2 = re.findall(
                                r'(' + pron + r')\W+(?:\w+\W+){0,' + str(mxpron) + r'}?(' + stext + r')',
                                filtered_sent, re.IGNORECASE)
                            k = k1 + k2
                            k = ''.join(''.join(elems) for elems in k)
                            k = re.sub(r'(' + qs + r')', '', k)
                            if len(k) > 0 and no_prows == 0:
                                h.last = person_set(most_recent_person, next_person)[0]
                                h.person = person_set(most_recent_person, next_person)[1]
                                h.role = person_set(most_recent_person, next_person)[2]
                                h.firm = person_set(most_recent_person, next_person)[3]
                                h.former = person_set(most_recent_person, next_person)[4]
                                h.qsaid = "yes"
                                h.qpref = "no"
                                h.comment = "pronoun"
                                for q in quotes:
                                    h.qtype = "single sentence quote"
                                    h.quote = q
                                    h.todf()
                                    dx += 1
                                if multi_sent_indicator == 1 and len(msent_quote) > 0:
                                    for i in range(0, len(msent_quotes)):
                                        h.quote = msent_quotes[i]
                                        h.sentence = stences[i]
                                        h.sentiment_stanza = sentiments[i]
                                        h.sent_vader_neg = sent_vader_negs[i]
                                        h.sent_vader_neu = sent_vader_neus[i]
                                        h.sent_vader_pos = sent_vader_poss[i]
                                        h.n = nindex[i]
                                        h.qtype = "multi-sentence quote"
                                        h.todf()
                                        h.n = n

                            # If said equiv and quote, but no people or pronoun, check if organization listed,
                            # as in "BMW's spokesman said..."
                            elif len(k) == 0:
                                if len(orgs) == 1:
                                    # check if spokeperson
                                    org = orgs[0]
                                    k3 = re.findall(r'' + org + r'.+(' + rep + r')', filtered_sent, re.IGNORECASE)
                                    k4 = re.findall(r'' + rep + r'.+(' + org + r')', filtered_sent, re.IGNORECASE)
                                    k = k3 + k4
                                    k = ''.join(''.join(elems) for elems in k)
                                    k = re.sub(r'(' + qs + r')', '', k)
                                    if len(k) > 0:
                                        h.last = "NA"
                                        # The person would be the representative's reference,
                                        # removing the org name
                                        h.person = "NA"
                                        h.firm = org
                                        h.role = re.sub(r'' + org + '', '', k)
                                        h.former = former_search(filtered_sent, h.role)
                                        h.qsaid = "yes"
                                        h.qpref = "no"
                                        h.comment = "Org representative"
                                        for q in quotes:
                                            h.quote = q
                                            h.qtype = "single sentence quote"
                                            h.todf()
                                            dx += 1
                                        if multi_sent_indicator == 1 and len(msent_quote) > 0:
                                            for i in range(0, len(msent_quotes)):
                                                h.quote = msent_quotes[i]
                                                h.sentence = stences[i]
                                                h.sentiment_stanza= sentiments[i]
                                                h.sent_vader_neg = sent_vader_negs[i]
                                                h.sent_vader_neu = sent_vader_neus[i]
                                                h.sent_vader_pos = sent_vader_poss[i]
                                                h.n = nindex[i]
                                                h.qtype = "multi-sentence quote"
                                                h.todf()
                                                h.n = n

                                    elif len(k) == 0:
                                        h.last = "NA"
                                        h.person = "NA"
                                        h.firm = org
                                        h.role = "NA"
                                        h.former = "NA"
                                        h.qsaid = "yes"
                                        h.qpref = "no"
                                        h.comment = "from org, nondiscriptives"
                                        for q in quotes:
                                            h.quote = q
                                            h.qtype = "single sentence quote"
                                            h.todf()
                                            dx += 1
                                        if multi_sent_indicator == 1 and len(msent_quote) > 0:
                                            for i in range(0, len(msent_quotes)):
                                                h.quote = msent_quotes[i]
                                                h.sentence = stences[i]
                                                h.sentiment_stanza= sentiments[i]
                                                h.sent_vader_neg = sent_vader_negs[i]
                                                h.sent_vader_neu = sent_vader_neus[i]
                                                h.sent_vader_pos = sent_vader_poss[i]
                                                h.n = nindex[i]
                                                h.qtype = "multi-sentence quote"
                                                h.todf()
                                                h.n = n

                                elif len(orgs) > 1:
                                    for org in orgs:
                                        # check if spokeperson
                                        k1 = re.findall(
                                            r'(' + rep + r')\W+(?:\w+\W+){0,' + str(mxorg1) + r'}?(' + org + r')',
                                            filtered_sent, re.IGNORECASE)
                                        k2 = re.findall(
                                            r'(' + org + r')\W+(?:\w+\W+){0,' + str(mxorg1) + r'}?(' + rep + r')',
                                            filtered_sent, re.IGNORECASE)
                                        k = k1 + k2
                                        k = ''.join(''.join(elems) for elems in k)
                                        k = re.sub(r'(' + qs + r')', '', k)
                                        if len(k) > 0:
                                            h.last = "NA"
                                            # The person would be the representative's reference,
                                            # removing the org name
                                            h.person = "NA"
                                            h.firm = org
                                            h.role = re.sub(r'' + org + '', '', k)
                                            h.former = former_search(filtered_sent, h.role)
                                            h.qsaid = "yes"
                                            h.qpref = "no"
                                            h.comment = "Org representative"
                                            for q in quotes:
                                                h.quote = q
                                                h.qtype = "single sentence quote"
                                                h.todf()
                                                dx += 1
                                            if multi_sent_indicator == 1 and len(msent_quote) > 0:
                                                for i in range(0, len(msent_quotes)):
                                                    h.quote = msent_quotes[i]
                                                    h.sentence = stences[i]
                                                    h.sentiment_stanza= sentiments[i]
                                                    h.sent_vader_neg = sent_vader_negs[i]
                                                    h.sent_vader_neu = sent_vader_neus[i]
                                                    h.sent_vader_pos = sent_vader_poss[i]
                                                    h.n = nindex[i]
                                                    h.qtype = "multi-sentence quote"
                                                    h.todf()
                                                    h.n = n
                                            break

                                        elif len(k) == 0:
                                            h.last = "NA"
                                            h.person = "NA"
                                            h.firm = ''.join(''.join(elems) for elems in orgs)
                                            h.role = "NA"
                                            h.former = "NA"
                                            h.qtype = "Information - non-discript single sentence quote"
                                            h.qsaid = "yes"
                                            h.qpref = "no"
                                            h.comment = "Review - from nondiscriptives"
                                            for q in quotes:
                                                h.quote = q
                                                h.qtype = "single sentence quote"
                                                h.todf()
                                                dx += 1
                                            if multi_sent_indicator == 1 and len(msent_quote) > 0:
                                                for i in range(0, len(msent_quotes)):
                                                    h.quote = msent_quotes[i]
                                                    h.sentence = stences[i]
                                                    h.sentiment_stanza= sentiments[i]
                                                    h.sent_vader_neg = sent_vader_negs[i]
                                                    h.sent_vader_neu = sent_vader_neus[i]
                                                    h.sent_vader_pos = sent_vader_poss[i]
                                                    h.n = nindex[i]
                                                    h.qtype = "multi-sentence quote"
                                                    h.todf()
                                                    h.n = n



                                elif len(orgs) == 0:
                                    # if there is a said equivalent and quote but no person, no org, and no pronoun
                                    # near said equivalent, check if pronoun further back from said equiv.
                                    k1 = re.findall(
                                        r'(' + stext + r').+(' + pron + r').+(' + start + r').+(' + end + r')',
                                        stence, re.IGNORECASE)
                                    k2 = re.findall(
                                        r'(' + pron + r').+(' + stext + r').+(' + start + r').+(' + end + r')',
                                        stence, re.IGNORECASE)
                                    k3 = re.findall(
                                        r'(' + start + r').+(' + end + r').+(' + pron + r').+(' + stext + r')',
                                        stence, re.IGNORECASE)
                                    k4 = re.findall(
                                        r'(' + start + r').+(' + end + r').+(' + stext + r').+(' + pron + r')',
                                        stence, re.IGNORECASE)

                                    k = k1 + k2 + k3 + k4
                                    k = ''.join(''.join(elems) for elems in k)
                                    k = re.sub(r'(' + qs + r')', '', k)
                                    # If there was a pronoun and said equiv outside of quotes, establish who
                                    # the reference is
                                    if len(k) > 0:

                                        # First, check when the last person reference was
                                        # The second items, the item in the person list, must be last because
                                        # sppl has n ongoing
                                        # But sppl may not exist yet so have condition here.
                                        try:
                                            m = int(sppl[-1][-1])
                                            l = n - int(m)
                                        except:
                                            l = 2
                                        # If person in the previous sentence, use that one
                                        if l == 1 and no_prows == 0:
                                            h.last = person_set(most_recent_person, next_person)[0]
                                            h.person = person_set(most_recent_person, next_person)[1]
                                            h.role = person_set(most_recent_person, next_person)[2]
                                            h.firm = person_set(most_recent_person, next_person)[3]
                                            h.former = person_set(most_recent_person, next_person)[4]
                                            h.qsaid = "yes"
                                            h.qpref = "no"
                                            h.comment = "pronoun"
                                            for q in quotes:
                                                h.quote = q
                                                h.qtype = "single sentence quote"
                                                h.todf()
                                                dx += 1
                                            if multi_sent_indicator == 1 and len(msent_quote) > 0:
                                                for i in range(0, len(msent_quotes)):
                                                    h.quote = msent_quotes[i]
                                                    h.sentence = stences[i]
                                                    h.sentiment_stanza= sentiments[i]
                                                    h.sent_vader_neg = sent_vader_negs[i]
                                                    h.sent_vader_neu = sent_vader_neus[i]
                                                    h.sent_vader_pos = sent_vader_poss[i]
                                                    h.n = nindex[i]
                                                    h.qtype = "multi-sentence quote"
                                                    h.todf()
                                                    h.n = n
                                        # if no person in previous sentence, then check if quote
                                        # in previous sentence
                                        elif l > 1:
                                            # Check if there's something in DF:
                                            if dx>0:
                                                s = df.loc[len(df.index) - 1].at['sent']
                                                t = n - int(s)
                                                if t == 1:
                                                    h.last = df.loc[len(df.index) - 1].at['last']
                                                    h.person = df.loc[len(df.index) - 1].at['person']
                                                    h.firm = df.loc[len(df.index) - 1].at['firm']
                                                    h.role = df.loc[len(df.index) - 1].at['role']
                                                    h.former = df.loc[len(df.index) - 1].at['former']
                                                    h.qsaid = "yes"
                                                    h.qpref = "no"
                                                    h.comment = "pronoun - reference partial profile"
                                                    for q in quotes:
                                                        h.quote = q
                                                        h.qtype = "single sentence quote"
                                                        h.todf()
                                                        dx += 1
                                                    if multi_sent_indicator == 1 and len(msent_quote) > 0:
                                                        for i in range(0, len(msent_quotes)):
                                                            h.quote = msent_quotes[i]
                                                            h.sentence = stences[i]
                                                            h.sentiment_stanza= sentiments[i]
                                                            h.sent_vader_neg = sent_vader_negs[i]
                                                            h.sent_vader_neu = sent_vader_neus[i]
                                                            h.sent_vader_pos = sent_vader_poss[i]
                                                            h.n = nindex[i]
                                                            h.qtype = "multi-sentence quote"
                                                            h.todf()
                                                            h.n = n

                        # If said equivalent, no people and no quote, check for plurals
                        elif len(x1) == 0:
                            m = re.findall(r'(' + plurals + r')', filtered_sent, re.IGNORECASE)

                            if len(m) > 0:
                                h.last = "NA"
                                h.person = "NA"
                                h.firm = "NA"
                                h.role = ''.join(''.join(elems) for elems in m)
                                h.former = former_search(filtered_sent, h.role)
                                h.quote = stence
                                h.qtype = "paraphrase"
                                h.qsaid = "yes"
                                h.qpref = "no"
                                h.comment = "from plural individuals"
                                h.todf()
                                dx += 1


                            # If said equivalent, no people, no quote, no plural, check if pronoun, rep,
                            # and finally Org
                            elif len(m) == 0:

                                k1 = re.findall(
                                    r'(' + stext + r')\W+(?:\w+\W+){0,' + str(mxpron) + r'}?(' + pron + r')',
                                    filtered_sent, re.IGNORECASE)
                                k2 = re.findall(
                                    r'(' + pron + r')\W+(?:\w+\W+){0,' + str(mxpron) + r'}?(' + stext + r')',
                                    filtered_sent, re.IGNORECASE)
                                k = k1 + k2
                                k = ''.join(''.join(elems) for elems in k)
                                k = re.sub(r'(' + qs + r')', '', k)
                                if len(k) > 0 and no_prows == 0:
                                    h.quote = stence
                                    h.last = person_set(most_recent_person, next_person)[0]
                                    h.person = person_set(most_recent_person, next_person)[1]
                                    h.role = person_set(most_recent_person, next_person)[2]
                                    h.firm = person_set(most_recent_person, next_person)[3]
                                    h.former = person_set(most_recent_person, next_person)[4]
                                    h.qtype = "paraphrase"
                                    h.qsaid = "yes"
                                    h.qpref = "no"
                                    h.comment = "pronoun"
                                    h.todf()
                                    dx += 1


                                elif len(k) == 0:
                                    # If said equiv, but no person, no quote, no plural, and no pronoun, check if
                                    # end of quote, in which case you would just pass. This has already been captured.
                                    q2a = re.findall(r'(' + end + r')(.+)', stence, re.IGNORECASE)
                                    q2a = ''.join(''.join(elems) for elems in q2a)
                                    if len(q2a) > 0:
                                        pass
                                    else:
                                        # If not end quote either, check if a title such as "The executive said..."
                                        role_list = rolesearch(stence)
                                        role_list = ''.join(''.join(elems) for elems in role_list)
                                        if len(role_list) > 0 and no_prows == 0:
                                            h.quote = stence
                                            h.last = person_set(most_recent_person, next_person)[0]
                                            h.person = person_set(most_recent_person, next_person)[1]
                                            h.role = person_set(most_recent_person, next_person)[2]
                                            h.firm = person_set(most_recent_person, next_person)[3]
                                            h.former = person_set(most_recent_person, next_person)[4]
                                            h.qtype = "paraphrase"
                                            h.qsaid = "yes"
                                            h.qpref = "no"
                                            h.comment = "title reference"
                                            h.todf()
                                            dx += 1
                                        elif len(role_list) == 0:
                                            # see if it's a representative of a firm
                                            if len(orgs) > 0:
                                                for org in orgs:
                                                    m1 = re.findall(r'' + org + r'\s+(' + rep + r')', stence,
                                                                    re.IGNORECASE)
                                                    m2 = re.findall(r'(' + rep + r')\s+of\s+' + org + r'',
                                                                    stence,
                                                                    re.IGNORECASE)
                                                    m = m1 + m2
                                                    m = ''.join(''.join(elems) for elems in m)
                                                    m = re.sub(r'(' + org + r')', '', m)
                                                    if len(m) > 0:
                                                        h.last = "NA"
                                                        h.person = "NA"
                                                        h.firm = org
                                                        h.role = m
                                                        h.former = former_search(filtered_sent, h.role)
                                                        h.qtype = "paraphrase"
                                                        h.qsaid = "yes"
                                                        h.qpref = "one or more firms"
                                                        h.comment = "a firm representative said something"
                                                        h.quote = stence
                                                        h.todf()
                                                        dx += 1
                                                    # If not a rep, see if just one company
                                                    # If so, the company said something
                                                    elif len(m) == 0:
                                                        if len(orgs) == 1:
                                                            h.quote = stence
                                                            h.last = "NA"
                                                            h.person = "NA"
                                                            h.firm = org
                                                            h.role = "NA"
                                                            h.former = "NA"
                                                            h.qtype = "paraphrase"
                                                            h.qsaid = "yes"
                                                            h.qpref = "no"
                                                            h.comment = 'One firm said something'
                                                            h.todf()
                                                            dx += 1
                                                        # If not, see if multiple companies
                                                        # If so, add line for each org and said, but flag to review
                                                        elif len(orgs) > 1:
                                                            h.quote = stence
                                                            h.last = "NA"
                                                            h.person = "NA"
                                                            h.firm = org
                                                            h.role = "NA"
                                                            h.former = "NA"
                                                            h.qtype = "paraphrase"
                                                            h.qsaid = "yes"
                                                            h.qpref = "no"
                                                            h.comment = "Review - paraphrase from multiple firms"
                                                            h.todf()
                                                            dx += 1


                        # Because said equivalent but no full quote, check if it's a beginning of
                        # multi-sentence quote
                        else:
                            # If there is a multi-sentence quote, then do stuff
                            if msent_quote is not None:
                                if len(msent_quote) > 0:
                                    # We know there is a said equivalent in the first sentence
                                    # So now it's just checking who said it.
                                    # First, check if person:
                                    if len (first_ppl) > 0:
                                        if len(first_pron) == 0 and no_prows == 0:
                                            for i in range(0, len(msent_quotes)):

                                                h.quote = msent_quotes[i]
                                                h.sentence = stences[i]
                                                h.sentiment_stanza= sentiments[i]
                                                h.sent_vader_neg = sent_vader_negs[i]
                                                h.sent_vader_neu = sent_vader_neus[i]
                                                h.sent_vader_pos = sent_vader_poss[i]
                                                h.n = nindex[i]
                                                # Then it must be the most recent person, who is from this sentence
                                                h.last = person_set(most_recent_person, next_person)[0]
                                                h.person = person_set(most_recent_person, next_person)[1]
                                                h.role = person_set(most_recent_person, next_person)[2]
                                                h.firm = person_set(most_recent_person, next_person)[3]
                                                h.former = person_set(most_recent_person, next_person)[4]
                                                h.qtype = "multi-sentence quote"
                                                h.qsaid = "yes"
                                                h.qpref = "no"
                                                h.comment = "Person named in first sentence"
                                                h.todf()
                                                h.n = n

                                        elif len(first_pron) > 0 and no_prows == 0:
                                            for i in range(0, len(msent_quotes)):
                                                h.quote = msent_quotes[i]
                                                h.sentence = stences[i]
                                                # Then it must be the most recent person
                                                h.last = person_set(most_recent_person, next_person)[0]
                                                h.person = person_set(most_recent_person, next_person)[1]
                                                h.role = person_set(most_recent_person, next_person)[2]
                                                h.firm = person_set(most_recent_person, next_person)[3]
                                                h.former = person_set(most_recent_person, next_person)[4]
                                                h.qtype = "multi-sentence quote"
                                                h.qsaid = "yes"
                                                h.qpref = "no"
                                                h.comment = "Pronoun"
                                                h.todf()


                    # if said something and exactly one person from list, do stuff
                    elif len(ppl) == 1 and len(prows)>0:

                        ##################################
                        # ONE AND ONLY ONE PEOPLE REFERENCE
                        ##################################
                        h.last = last = prows[0][0]

                        # Check for FULL quotes:
                        # first check for the FULL quote after the said equivalent
                        # last name then said equivalent then quote
                        s1a = re.findall(r'(' + last + r').+(' + stext + r').+(' + start + r')(.+)(' + end + r')',
                                         stence,
                                         re.IGNORECASE)
                        # said equiv then last name then a quote
                        s1b = re.findall(r'(' + stext + r').+(' + last + r').+(' + start + r')(.+)(' + end + r')',
                                         stence,
                                         re.IGNORECASE)
                        # quote then said-equivalent then last name
                        s2a = re.findall(r'(' + start + r')(.+)(' + end + r').+(' + stext + r').+(' + last + r')',
                                         stence,
                                         re.IGNORECASE)
                        # quote then last name then said-equivalent
                        s2b = re.findall(r'(' + start + r')(.+)(' + end + r').+(' + last + r').+(' + stext + r')',
                                         stence,
                                         re.IGNORECASE)
                        s1 = s1a + s1b + s2a + s2b
                        s1 = ''.join(''.join(elems) for elems in s1)
                        s1 = re.sub(r'(' + qs + r')', '', s1)
                        s1 = re.sub(r'(' + stext + r')', '', s1)
                        s1 = re.sub(r'' + last + r'', '', s1)

                        if len(s1) > 0:
                            h.last = prows[0][0]
                            h.person = prows[0][1]
                            h.role = prows[0][2]
                            h.firm = prows[0][3]
                            h.former = prows[0][4]
                            h.qsaid = "yes"
                            h.qpref = "single"
                            h.comment = "high accuracy"
                            for q in quotes:
                                h.quote = q
                                h.qtype = "single sentence quote"
                                h.todf()
                                dx += 1
                            if multi_sent_indicator == 1:
                                if len(msent_quote) > 0:
                                    for i in range(0, len(msent_quotes)):
                                        h.quote = msent_quotes[i]
                                        h.sentence = stences[i]
                                        h.sentiment_stanza= sentiments[i]
                                        h.sent_vader_neg = sent_vader_negs[i]
                                        h.sent_vader_neu = sent_vader_neus[i]
                                        h.sent_vader_pos = sent_vader_poss[i]
                                        h.n = nindex[i]
                                        h.qtype = "multi-sentence quote"
                                        h.todf()
                                        h.n = n


                        # Check if the multiple sentence quote
                        elif len(s1) == 0:

                            if msent_quote is not None:
                                if len(msent_quote) > 0:
                                    for i in range(0, len(msent_quotes)):
                                        h.quote = msent_quotes[i]
                                        h.sentence = stences[i]
                                        h.sentiment_stanza= sentiments[i]
                                        h.sent_vader_neg = sent_vader_negs[i]
                                        h.sent_vader_neu = sent_vader_neus[i]
                                        h.sent_vader_pos = sent_vader_poss[i]
                                        h.n = nindex[i]
                                        h.last = prows[0][0]
                                        h.person = prows[0][1]
                                        h.role = prows[0][2]
                                        h.firm = prows[0][3]
                                        h.former = prows[0][4]
                                        h.qtype = "multi-sentence quote"
                                        h.qsaid = "yes"
                                        h.qpref = "single"
                                        h.comment = "high accuracy"
                                        h.todf()
                                        h.n = n
                                        dx += 1
                            # If there is no full quote and no multi-sentence quote, but there is an ending quote and a person,
                            # update the previous row with this person. The previous row was a multi-sentence quote
                            # defaulted as the most previous person mentioned. So this will just confirm the previous person,
                            # or update it appropriately
                            q1a = re.findall(r'(' + end + r').+(' + stext + r').+(' + last + r')', stence,
                                             re.IGNORECASE)
                            q1b = re.findall(r'(' + end + r').+(' + last + r').+(' + stext + r')', stence,
                                             re.IGNORECASE)
                            q1 = q1a + q1b
                            q1 = ''.join(''.join(elems) for elems in q1)
                            q = re.sub(r'(' + qs + r')', '', q1)
                            # Just update the previous person profile
                            if len(q) > 0:
                                df.loc[len(df.index) - 1].at['last'] = prows[0][0]
                                df.loc[len(df.index) - 1].at['person'] = prows[0][1]
                                df.loc[len(df.index) - 1].at['firm'] = prows[0][2]
                                df.loc[len(df.index) - 1].at['role'] = prows[0][3]
                                df.loc[len(df.index) - 1].at['former'] = prows[0][4]

                            elif len(q) == 0:
                                h.quote = stence
                                h.last = prows[0][0]
                                h.person = prows[0][1]
                                h.role = prows[0][2]
                                h.firm = prows[0][3]
                                h.former = prows[0][4]
                                h.qtype = "paraphrase"
                                h.qsaid = "yes"
                                h.qpref = "single"
                                h.comment = "high accuracy"
                                h.todf()
                                dx += 1


                    # If said equivalent but more than one person, do stuff:
                    elif len(ppl) > 1 and len(prows)>0:
                        ##################################
                        # MORE THAN ONE PEOPLE REFERENCED
                        ##################################
                        for per in prows:
                            h.last = per[0]
                            h.person = per[1]
                            h.role = per[2]
                            h.firm = per[3]
                            h.former = per[4]

                            # Check if full quote
                            # Quote then said-equivalent is within three spaces of last name
                            s1a = re.findall(
                                r'(' + start + r')(.+)(' + end + r')(?=.+(' + h.last + r')\W+(?:\w+\W+){0,' + str(
                                    mxlds) + r'}?' + stext + r')', stence, re.IGNORECASE)
                            # Said equiv then last name then a quote (distance not matter here because sometimes full title in between)
                            s1b = re.findall(
                                r'(' + h.last + r').+(' + stext + 'r).+(' + start + r')(.+)(' + end + r')',
                                stence, re.IGNORECASE)
                            s1 = s1a + s1b
                            s1 = ''.join(''.join(elems) for elems in s1)
                            s1 = re.sub(r'(' + qs + r')', '', s1)
                            s1 = re.sub(r'(' + stext + r')', '', s1)
                            s1 = re.sub(r'' + h.last + r'', '', s1)
                            if len(s1) > 0:
                                quote = s1
                                quote = re.sub(r'^' + stext + r'.+', '', quote)
                                h.quote = re.sub(r'' + stext + r'\Z', '', quote)
                                h.qtype = "single sentence quote"
                                h.qsaid = "yes"
                                h.qpref = "multiple"
                                h.comment = "review"
                                h.todf()
                                dx += 1
                            # check if a quote begins in this sentence:
                            elif len(s1) == 0:
                                q1a = re.findall(r'(' + h.last + r').+(' + stext + r').+(' + start + r')(.+)', stence,
                                                 re.IGNORECASE)
                                # said equiv then last name then a quote
                                q1b = re.findall(r'(' + stext + r').+(' + h.last + r').+(' + start + r')(.+)', stence,
                                                 re.IGNORECASE)
                                q1 = q1a + q1b
                                q1 = ''.join(''.join(elems) for elems in q1)
                                q1 = re.sub(r'(' + qs + r')', '', q1)
                                q1 = re.sub(r'(' + stext + r')', '', q1)
                                q1 = re.sub(r'' + h.last + r'', '', q1)

                                # If a quote does begin in this sentence, concatinate the subsequent sentences until the quote ends
                                if len (q1)>0:
                                    if msent_quote is not None:
                                        if len(msent_quote) > 0:
                                            for i in range(0, len(msent_quotes)):
                                                h.quote = msent_quotes[i]
                                                h.sentence = stences[i]
                                                h.sentiment_stanza= sentiments[i]
                                                h.sent_vader_neg = sent_vader_negs[i]
                                                h.sent_vader_neu = sent_vader_neus[i]
                                                h.sent_vader_pos = sent_vader_poss[i]
                                                h.n = nindex[i]
                                                h.last = prows[0][0]
                                                h.person = prows[0][1]
                                                h.role = prows[0][2]
                                                h.firm = prows[0][3]
                                                h.former = prows[0][4]
                                                h.qtype = "multi-sentence quote"
                                                h.qsaid = "yes"
                                                h.qpref = "multiple"
                                                h.comment = "review"
                                                h.todf()
                                                h.n = n
                                                dx += 1

                                elif len(q1) == 0:
                                    h.quote = stence
                                    h.qtype = "paraphrase"
                                    h.qsaid = "yes"
                                    h.qpref = "multiple"
                                    h.comment = "review"
                                    h.todf()
                                    dx += 1


                elif len(s) == 0:
                ##################################
                # SENTENCES WITH NO "SAID" EQUIVALENT
                ##################################
                    # Check if no person:
                    if len(ppl) == 0:
                        ##################################
                        # NO PEOPLE REFERENCE
                        ##################################
                        # Since no clear person reference, check if 2+ words in a quote:
                        x1 = re.findall(r'(' + start + r')(\b.+\b\s+\b.+\b.*)(' + end + r')', stence, re.IGNORECASE)
                        x1 = ''.join(''.join(elems) for elems in x1)
                        x1 = re.sub(r'(' + qs + r')', '', x1)
                        # if so, must be from the most recent person referenced in the article

                        if len(x1) > 0 and no_prows == 0:
                            h.last = person_set(most_recent_person, next_person)[0]
                            h.person = person_set(most_recent_person, next_person)[1]
                            h.role = person_set(most_recent_person, next_person)[2]
                            h.firm = person_set(most_recent_person, next_person)[3]
                            h.former = person_set(most_recent_person, next_person)[4]
                            h.qsaid = "no"
                            h.qpref = "no"
                            h.comment = "Quote with at least two words inside"
                            for q in quotes:
                                h.quote = q
                                h.qtype = "single sentence quote"
                                h.todf()
                                dx += 1
                            if msent_quote is not None:
                                if len(msent_quote) > 0:
                                    for i in range(0, len(msent_quotes)):
                                        h.quote = msent_quotes[i]
                                        h.sentence = stences[i]
                                        h.sentiment_stanza= sentiments[i]
                                        h.sent_vader_neg = sent_vader_negs[i]
                                        h.sent_vader_neu = sent_vader_neus[i]
                                        h.sent_vader_pos = sent_vader_poss[i]
                                        h.n = nindex[i]
                                        h.qtype = "multi-sentence quote"
                                        h.todf()
                                        h.n = n


                        # Because no said equiv or full quote, check if it's a beginning of multi-sentence quote
                        elif len(x1) == 0:

                            # Check if just one word quoted, keep going and do nothing
                            g = re.findall(r'(' + start + r')(\b.+\b)(' + end + r')', stence, re.IGNORECASE)
                            g = ''.join(''.join(elems) for elems in g)
                            if len(g) > 0:
                                pass
                            # If not, check if a quote begins in this sentence:
                            else:
                                # If there is a multi-sentence quote, then do stuff
                                if msent_quote is not None:
                                    if len(msent_quote) > 0:
                                        # Check if there's a person in that last sentence
                                        if len(last_ppl) > 0 and no_people == 0 and len(prows)>0:
                                            for i in range(0, len(msent_quotes)):
                                                h.quote = msent_quotes[i]
                                                h.sentence = stences[i]
                                                h.sentiment_stanza= sentiments[i]
                                                h.sent_vader_neg = sent_vader_negs[i]
                                                h.sent_vader_neu = sent_vader_neus[i]
                                                h.sent_vader_pos = sent_vader_poss[i]
                                                h.n = nindex[i]
                                                h.last = dfsent.loc[last_sindex].at['prows'][0][0]
                                                h.person = dfsent.loc[last_sindex].at['prows'][0][1]
                                                h.role = dfsent.loc[last_sindex].at['prows'][0][2]
                                                h.firm = dfsent.loc[last_sindex].at['prows'][0][3]
                                                h.former = dfsent.loc[last_sindex].at['prows'][0][4]
                                                h.qtype = "multi-sentence quote"
                                                h.qsaid = "no"
                                                h.qpref = "no"
                                                h.comment = "high accuracy"
                                                h.todf()
                                                h.n = n
                                                dx += 1

                                        elif len(last_ppl) == 0 and no_prows == 0:
                                            for i in range(0, len(msent_quotes)):
                                                h.quote = msent_quotes[i]
                                                h.sentence = stences[i]
                                                h.sentiment_stanza= sentiments[i]
                                                h.sent_vader_neg = sent_vader_negs[i]
                                                h.sent_vader_neu = sent_vader_neus[i]
                                                h.sent_vader_pos = sent_vader_poss[i]
                                                h.n = nindex[i]
                                                h.last = person_set(most_recent_person, next_person)[0]
                                                h.person = person_set(most_recent_person, next_person)[1]
                                                h.role = person_set(most_recent_person, next_person)[2]
                                                h.firm = person_set(most_recent_person, next_person)[3]
                                                h.former = person_set(most_recent_person, next_person)[4]
                                                h.qtype = "multi-sentence quote"
                                                h.qsaid = "no"
                                                h.qpref = "no"
                                                h.comment = "high accuracy"
                                                h.todf()
                                                h.n = n
                                                dx += 1


                    ##################################
                    # ONE OR MORE PEOPLE REFERENCED
                    ##################################
                    elif len(ppl) > 0 and len(prows)>0:
                        # Since no said equivalent, but one or more people referenced,
                        # check for quotes because if quoted then the most recent person referenced should be the speaker and the person/people
                        # referenced in the sentence would either be inside a quote as in '"It was such a great year for Steve Jobs."')
                        # or just in the sentence as in "'We caught up with him and he shared his thoughts about Steve Jobs: "What a great year he had."

                        x1 = re.findall(r'(' + start + r')(.+)(' + end + r')', stence, re.IGNORECASE)
                        x1 = ''.join(''.join(elems) for elems in x1)
                        x1 = re.sub(r'(' + qs + r')', '', x1)

                        # if so, assume must be from the most recent person referenced in the article
                        if len(x1) > 0:
                            d = []
                            # check if quote is within "person"
                            for p in fulls:
                                t = re.findall(r'(' + start + r')(.+)(' + end + r')', p, re.IGNORECASE)
                                if len(t) > 0:
                                    d.append(t)
                            if len(d) == 0:
                                h.last = "NA"
                                h.person = "NA"
                                h.firm = "NA"
                                h.role = "NA"
                                h.former = "NA"
                                h.qsaid = "no"
                                h.qpref = "one or more"
                                h.comment = "Review, whether quote or something else, not clear"
                                for q in quotes:
                                    h.quote = q
                                    h.qtype = "single sentence quote"
                                    h.todf()
                                    dx += 1
                                if msent_quote is not None:
                                    if multi_sent_indicator == 1:
                                        for i in range(0, len(msent_quotes)):
                                            h.quote = msent_quotes[i]
                                            h.sentence = stences[i]
                                            h.sentiment_stanza= sentiments[i]
                                            h.sent_vader_neg = sent_vader_negs[i]
                                            h.sent_vader_neu = sent_vader_neus[i]
                                            h.sent_vader_pos = sent_vader_poss[i]
                                            h.n = nindex[i]
                                            h.qtype = "multi-sentence quote"
                                            h.todf()
                                            h.n = n
                        # Because said equivalent but no full quote, check if it's a beginning of multi-sentence quote
                        # If so, assume it would have to be from the most recent referenced person
                        # As in "he said", referring to the most recent person
                        elif len(x1) == 0:
                            # Check if multi-sentence
                            if msent_quote is not None:
                                if len(msent_quote) > 0 and no_prows == 0:
                                    for i in range(0, len(msent_quotes)):
                                        h.quote = msent_quotes[i]
                                        h.sentence = stences[i]
                                        h.sentiment_stanza= sentiments[i]
                                        h.sent_vader_neg = sent_vader_negs[i]
                                        h.sent_vader_neu = sent_vader_neus[i]
                                        h.sent_vader_pos = sent_vader_poss[i]
                                        h.n = nindex[i]
                                        h.last = person_set(most_recent_person, next_person)[0]
                                        h.person = person_set(most_recent_person, next_person)[1]
                                        h.role = person_set(most_recent_person, next_person)[2]
                                        h.firm = person_set(most_recent_person, next_person)[3]
                                        h.former = person_set(most_recent_person, next_person)[4]
                                        h.qtype = "multi-sentence quote"
                                        h.qsaid = "no"
                                        h.qpref = "one or more"
                                        h.comment = "high accuracy - prior person referenced"
                                        h.todf()
                                        h.n = n
                                        dx += 1


                n += 1
                h.n += 1



        #except:
        #   df.loc[len(df.index)] = [n, "Issue", "Issue", "Issue", "Issue", "Issue", "Issue", "Issue", "Issue", "Issue",
        #                            "Issue", "Sentence-level Issue", AN, date, source, regions, subjects, misc, industries,
        #                            focfirm, by, comp, "Issue", "Issue", "Issue", "Issue"]

        t2 = time.time()
        total = t1 - t0
        print(art + 1, h.AN, "time run:", round(t2 - t1, 2), "total hours:", round((t2 - t0) / (60 * 60), 2),
              "mean rate:", round((t2 - t0) / (art + 1), 2))

    #except:
    #   df.loc[len(df.index)] = [n, "Issue", "Issue", "Issue", "Issue", "Issue", "Issue", "Issue", "Issue", "Issue",
    #                            "Issue", "Article NLP issue", AN, date, source, "Issue", "Issue", "Issue", "Issue",
    #                            "Issue", "Issue", "Issue", "Issue", "Issue", "Issue", "Issue"]




# output the chunk
# OUTFILE = f"{OUTDIR}/{INFILE.split('.')[0]}_df_quotes_{STARTROW.zfill(6)}_{ENDROW.zfill(6)}.xlsx"
# df.to_excel(OUTFILE)

df
#df.to_excel(r'C:\Users\danwilde\Dropbox (Penn)\Dissertation\Factiva\final_63_05_26.xlsx')
