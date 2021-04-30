import time
import sys
import os
import pandas as pd
import numpy as np
import stanza
import openpyxl
import re

# stanza.download('en')
nlp = stanza.Pipeline('en')

dfa = pd.read_excel(r'C:\Users\danwilde\Dropbox (Penn)\Dissertation\Factiva\final4.xlsx')

# combine the lead paragraph and the body of the article

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

columns = ['sent', 'firm', 'person', 'last', 'role',  'quote',  'qtype', 'qsaid', 'qpref', 'comment', 'AN', 'date',
           'source', 'regions', 'subjects', 'misc', 'industries', 'focfirm', 'by', 'comp', 'sentiment', 'sentence']

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
     r'sr.*director|managing\s+director|vice\s+minister|prime\s+minister|founding\s+partner|board\s+member' \
     r'business\s+owner|sales\s+manager|svp'

l4 = r'president|chair\S+|director|manager|vp|' \
     r'banker\b|specialist\b|accountant\b|journalist\b|reporter\b|analyst\b|consultant\b|negotiator\b|' \
     r'governor\b|congress\S+|house\s+speaker|senator|senior\s+fellow|fellow|undersecretary|' \
     r'spokesman\b|spokesperson\b|spokeswoman\b|representative\b|official\b|executive\b|\baide\b|\bpilot\b|professor|' \
     r'attorney|minister|secretary|\bguru\b|partner|general\s+counsel|mayor|dealer|board|administrator|owner|' \
     r'leader'

l5 = r'\bC.O\b'

pron = r'\bhe\b|\bshe\b'
plurals = r'analysts|bankers|consultants|executives|management|participants|officials'
rep = r'spokesman\b|spokesperson\b|spokeswoman\b|representative\b|official\b|executive\b'
start = r'\"|\'\'|\`\`'
end = r'\"|\'\''
qs = r'\`\`\"|\'\''
suf = r'\bJr\b|\bSr\b|\sI\s|\sII\s|\sIII\s|\bPhD\b|Ph\.D|\bMBA\b|\bCPA\b'
news = r'Reuters'

maxch = str(200)
maxch1 = str(60)


# Set the max number of words away from the focal leader and the focal organization
mxld = 6
mxorg = 6
mxlds = 2
mxpron = 2
mxorg1 = 2


# Create a class in which to place variables and then create a function that will place all the variables
# in a df row. Makes the code a bit more readable when you go through the todf function doezens of times
class Features:
    def __init__(self):
        self.df = self.qtype = self.qsaid = self.qpref = self.comment = self.n= self.person = \
        self.last = self.role = self.firm = self.quote = self.AN = self.date = self.source = \
        self.regions = self.subjects = self.misc = self.industries = self.focfirm = self.by = \
        self.comp = self.sentiment = self.stence = None

    def todf(self):
        df.loc[len(df.index)] = [self.n, self.firm, self.person, self.last, self.role,  self.quote, self.qtype,
                                           self.qsaid, self.qpref, self.comment, self.AN, self.date, self.source,
                                           self.regions, self.subjects, self.misc, self.industries, self.focfirm,
                                           self.by, self.comp, self.sentiment, self.stence]

        dfsent.at[n - 1, 'covered'] = 1

    def todf_info(self):
        df.loc[len(df.index)] = [self.n, self.firm, self.person, self.last, self.role,  self.quote, self.qtype,
                                           self.qsaid, self.qpref, self.comment, self.AN, self.date, self.source,
                                           self.regions, self.subjects, self.misc, self.industries, self.focfirm,
                                           self.by, self.comp, self.sentiment, self.stence]


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
    # Sentiment of current sentence, which is index n-1
    k = doc.sentences[n-1]
    sentiment = k.sentiment
    sentiments.append(sentiment)
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
            sentiment = j.sentiment
            sentiments.append(sentiment)
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
           last_said, last_pron, last_ppl, msent_quotes, stences, sentiments, nindex

# Function for when you want to search with instance of a role and another variable, when distance between
# doesn't matter.
def longrolesearch(sentence,non_role_var):
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
        r'\b(?:(' + non_role_var + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?(whose\s+(' + l1 + ')))',
        sentence,
        re.IGNORECASE)
    o3b = re.findall(
        r'\b(?:(' + non_role_var + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?(whose\s+(' + l2 + ')))',
        sentence,
        re.IGNORECASE)
    o3c = re.findall(
        r'\b(?:(' + non_role_var + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?(whose\s+(' + l3 + ')))',
        sentence,
        re.IGNORECASE)
    o3d = re.findall(
        r'\b(?:(' + non_role_var + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?(whose\s+(' + l4 + ')))',
        sentence,
        re.IGNORECASE)
    o3e = re.findall(
        r'\b(?:(' + non_role_var + r')\W+(?:\w+\W+){0,' + str(max_char) + r'}?(whose\s+(' + l5 + ')))',
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
        h.qtype = "information"
        h.qsaid = "NA"
        h.qpref = "NA"
        h.comment = "information about one firm"
        h.quote = stence
        h.todf_info()

    # Second, if multiple orgs, add that information
    else:
        for o in orgs:
            h.person = "NA"
            h.last = "NA"
            h.firm = o
            h.role = "NA"
            h.qtype = "information"
            h.qsaid = "NA"
            h.qpref = "NA"
            h.comment = "information about multiple firm"
            h.quote = stence
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
            h.qtype = "information"
            h.qsaid = "NA"
            h.qpref = "NA"
            h.comment = "information about person"
            h.quote = stence
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
    else:
        last = next_person[0]
        person = next_person[1]
        role = next_person[2]
        firm = next_person[3]
    return last, person, role, firm

z = 0

#arts = range(221,228)

arts = [230]
for art in arts:
#for art in dfa.index:
    #try:

        t1 = time.time()
        text = str(dfa['LP'][art]) + ' ' + str(dfa['TD'][art])
        #text = str(dfa['TD'][art])
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
        d3c = re.findall(r'^[A-Z]+\b.{0,' + maxch1 + r'}\b.+\b\s*--', text, re.IGNORECASE)

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
            text = re.sub(r'^[A-Z]+\b.{0,' + maxch1 + r'}\b.+\b\s*--', '', text)



        # Next, remove any instance of a set of non-white-space characters 1-letter parentheses (e.g., "(D)"). These mess up the NLP process as they mean something literally
        # in the PyTorch code.
        text = re.sub(r'\(\S+\)', '', text)
        # Next, remove any of several other characters.
        text = re.sub(r'[\[\]()#/]', '', text)
        # Next replace * with period. Another issue with nlp
        text = re.sub(r'\*', r'.', text)
        # Separate any words with multiple upper-case then lower case patterns (e.g., "MatrixOne" -> "Matrix One")
        text = re.sub(r"(\b[A-Z][a-z]+)([A-Z][a-z]+)", r"\1 \2", text)

        # Next remove any quotes around a single word (e.g., "great"). These also irratate the nlp program
        zy = []
        t = re.findall(r'(' + start + r')(\w+\w)(' + end + r')', text)
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


        ################################
        # Loading the text into the Stanza nlp system
        ################################
        doc = nlp(text)

        #try:
        people = []
        sppl = []

        fulls = []
        lsts = []


        n = 1
        h.n = 1
        #Set the sentence count to 0
        s = 0
        #Create a count variable for each article that starts at 0 and +1 every time you add to the DF
        dx = 0

        corgs = []

        # Create a sentence-level DFs
        columns = ['no.', 'stence', 'filtered_sent', 'sentiment', 'orgs', 'corgs', 'ppl', 'prows', 'people', 'covered']
        dfsent = pd.DataFrame(columns=columns)

        # Create a DF for people
        columns = ['first_sent', 'last','full','role','firm']
        dfpeople = pd.DataFrame(columns=columns)

        # Look through document and see if has instances of at least four all-cap words
        # broken up by non-word characters (e.g., spaces, commas) followed by a colon.
        f = re.findall(r'(((\b[A-Z]+\b\W*){4,}):)', text)
        comma_string = 0
        if len(f)> 0:
            comma_string = re.findall(r',', f[0][0])
        # If there is at least 1 of these very unique instances (with commas), this is got to be a transcript of some kind.
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
                    h.sentiment = sent.sentiment

                    #for ent in sent.ents:
                    #    print(n, ent.text, ent.type)

                    last = ""
                    role = ""
                    role1 = ""
                    person = ""
                    firm = ""
                    firm1 = ""
                    # Check if name title in the sentence:
                    g = re.findall(r'(((\b[A-Z]+\b\W*){4,}):)', stence)
                    if len(g) > 0:
                        a = g[0][0]
                        a = re.sub(r':', '', a)
                        # If company name is Inc and has a comma, remove the comma
                        a = re.sub(',\s+(INC|LLC)', '', a, re.IGNORECASE)
                        p = a.split(',')

                        person = p[0]
                        # Figure out the last name depending on if there's a suffix
                        sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                        if len(sfx) > 0:
                            last = person.split()[-2]
                        elif len(sfx) == 0:
                            last = person.split()[-1]
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
                            firm = role1
                        elif len(role_check)>0:
                            role = role1
                            firm = firm1

                        row = [last, person, role, firm, n]

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
                        # check if person in people list talking
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
                                    h.last = per[0]
                                    h.person = per[1]
                                    h.role = per[2]
                                    h.firm = per[3]
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
                                h.last = sppl[-1][0]
                                h.person = sppl[-1][1]
                                h.role = sppl[-1][2]
                                h.firm = sppl[-1][3]

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
                qtype = ""
                qsaid = ""
                qpref = ""
                comment = ""

                ppl = []
                lasts = []
                orgs = []
                prows = []
                prod = []

                #variables for future sentences as needed (for multi-sentence quotes)
                fppl = []
                flasts = []
                forgs = []
                fprows = []
                fprod = []

                h.stence = stence = sent.text


                # Calculate the sentiment for each sentence
                h.sentiment = sentiment = sent.sentiment

                #for ent in sent.ents:
                #    print(n, ent.text, ent.type)

                # Must stip the the sentence of any quotes in order to arrive at appropriate orgs, people, and products
                # First, remove any information within quotes:
                filtered_sent = re.sub(r'(' + start + r').+(' + end + r')', ' ', stence)
                # Second, remove information in beginning of multi-sentence quotes
                # the start quote would have a space before the quote
                filtered_sent = re.sub(r'([:|,])\s+(' + start + r')\S.+$', ' ', filtered_sent)
                # Third, remove cases in which the quote begins in the sentance and doesn't end.
                filtered_sent = re.sub(r'^(' + start + r')\S.+$', ' ', filtered_sent)
                # Last, remove information at the end of multi-sentence quotes
                # the end quote would have a space after the quote
                filtered_sent = re.sub(r'^.+(' + end + r')\s+', ' ', filtered_sent)


                # Check if you can add to dataset of people/firm/role
                # First, compile a list of people and organizations and products iteratively within the sentence.
                # You use products list in the orgs list so need to run them sep.
                for ent in sent.ents:
                    if ent.type == "PERSON":
                        z = re.findall(r'' + ent.text + r'', filtered_sent)
                        if len(z) > 0:
                            ppl.append(ent.text)
                            sfx = re.findall(r'(' + suf + r')', ent.text, re.IGNORECASE)
                            if len(sfx) > 0:
                                l = ent.text.split()[-2]
                                lasts.append(l)
                            elif len(sfx) == 0:
                                l = ent.text.split()[-1]
                                lasts.append(l)
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
                        tion = re.sub(r'[()-]', '', tion)
                        z = re.findall(r'' + tion + r'', filtered_sent)
                        if len(z) > 0:
                            # don't include acronyms
                            a = re.findall(r'\(\s*' + tion + r'\s*\)', stence)
                            a = ''.join(''.join(elems) for elems in a)
                            if len(a) == 0:
                                # Not an acronym, so see if any products in sentence:
                                # If not, see if already in list
                                if len(prod) == 0:
                                    if tion not in orgs:
                                        orgs.append(tion)
                                        corgs.append(tion)
                                # If not, check if it's not just a brand name with a product after it like Toyota SRV
                                else:
                                    d = []
                                    for p in prod:
                                        c = re.findall(r'' + ent.type + r'\s+' + p + r'', stence)
                                        d = d + c
                                    # If not, then just make sure it's not already in the list.
                                    # If not, append
                                    if len(d) == 0:
                                        if tion not in orgs:
                                            orgs.append(tion)
                                            corgs.append(tion)


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

                        if len(ppl) == 1:
                            # If exactly 1 org, not matter where the person is in proximity of title/role
                            if len(orgs) == 1:
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

                                # If yes, add as potential person
                                person = per
                                sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                                if len(sfx) > 0:
                                    last = person.split()[-2]
                                elif len(sfx) == 0:
                                    last = person.split()[-1]

                                role = role_set(role_list, per)

                            # Multiple orgs in sentence
                            elif len(orgs) > 1:

                                # Create list and variable that will be used at the end of this iteration of orgs
                                # "gm" checks if there are any instances in which an org is a substring of another
                                #  org name. If yes, len(g)>0. If not, len(g) = 0 and you're good.
                                c = 0
                                b = 0
                                q = 0
                                xl = []
                                op = ""
                                gh = []
                                gp = []
                                for o in orgs:
                                    or1 = orgs.copy()
                                    or1.remove(o)

                                    li = [match for match in or1 if o in match]
                                    gh.append(li)
                                    # If there is substring occurance, create a list with all firms with that substring
                                    if len(li) > 0:
                                        li.append(o)
                                        gp.append(li)
                                        gp = gp[0]

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

                                            os = shortrolesearch(filtered_sent, o, mxorg1)
                                            om = ''.join(''.join(elems) for elems in os)

                                            if len(os) > 0:
                                                person = per
                                                sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                                                if len(sfx) > 0:
                                                    last = person.split()[-2]
                                                elif len(sfx) == 0:
                                                    last = person.split()[-1]

                                                role = role_set(os, o)

                                                o = o.replace(r"'s", "")
                                                firm = o

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
                                                sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                                                if len(sfx) > 0:
                                                    last = person.split()[-2]
                                                elif len(sfx) == 0:
                                                    last = person.split()[-1]

                                                role = role_set(m, o)

                                                firm = o
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
                                                    sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                                                    if len(sfx) > 0:
                                                        last = person.split()[-2]
                                                    elif len(sfx) == 0:
                                                        last = person.split()[-1]

                                                    role = m[0][0]

                                                    firm = o
                                                    c += 1
                                                    q += 1
                                                    if len(gm) == 0:
                                                        break


                                                # If not any of these three strong matches ('s, of, whose),
                                                # check on a weaker match: if name are close to title
                                                elif len(j) == 0:

                                                    m = shortrolesearch(filtered_sent, o, mxorg1)
                                                    on = ''.join(''.join(elems) for elems in m)

                                                    # If so, check if person is also near the title
                                                    if len(on) > 0:
                                                        m = shortrolesearch(filtered_sent, per, mxorg1)
                                                        ol = ''.join(''.join(elems) for elems in m)


                                                        # If name and person are close to title, assign
                                                        if len(ol) > 0:
                                                            person = per
                                                            sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                                                            if len(sfx) > 0:
                                                                last = person.split()[-2]
                                                            elif len(sfx) == 0:
                                                                last = person.split()[-1]

                                                            role = role_set(m, per)

                                                            firm = o
                                                            c += 1
                                                            if len(gm) == 0:
                                                                break
                                                    # If not, check if name followed by just one org and a title.
                                                    elif len(on) == 0:
                                                        m = threerolesearch(filtered_sent, per, o)

                                                        if len(m) > 0:
                                                            for i in m:
                                                                j = ''.join(''.join(elems) for elems in i)
                                                                if len(j) > 0:
                                                                    xl.append(j)
                                                                    op = o
                                                            c += 1
                                                            if len(gm) == 0:
                                                                break

                                if len(xl) == 1:
                                    role = re.sub(r'' + op + r'', '', xl[0])
                                    firm = re.sub(r'' + role + r'', '', xl[0])
                                    person = per
                                    sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                                    if len(sfx) > 0:
                                        last = person.split()[-2]
                                    elif len(sfx) == 0:
                                        last = person.split()[-1]

                            # If there's a titled person with no org, it may be the org in previous reference
                            # If a repeat person, this will be filtered out below.
                            elif len(orgs) == 0:
                                if len(df) > 0:
                                    person = per
                                    sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                                    if len(sfx) > 0:
                                        last = person.split()[-2]
                                    elif len(sfx) == 0:
                                        last = person.split()[-1]
                                    role = role_set(role_list, per)
                                    firm = df.loc[len(df.index) - 1].at['firm']
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
                                            person = per
                                            sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                                            if len(sfx) > 0:
                                                last = person.split()[-2]
                                            elif len(sfx) == 0:
                                                last = person.split()[-1]

                        # Combine last name, person, firm, and role, for that person.
                        row = [last, person, role, firm, n]

                        # If something in each key varable then add to the dataset of people
                        if min(len(row[0]), len(row[1]), len(row[2]), len(row[3])) > 0:
                            # Filter out people who already were added to the list
                            if last in lsts:
                                pass
                            else:
                                people.append(row)
                                # remove duplicates
                                people = [list(x) for x in set(tuple(x) for x in people)]

                elif len(ppl) == 0:
                    if len(orgs) == 1:
                        o = orgs[0]
                        m = shortrolesearch(filtered_sent, o, mxorg1)

                        on = ''.join(''.join(elems) for elems in m)

                        # If there is a role and org close together, then check if a single person in the next sentence:
                        if len(on) > 0:
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

                                # Combine last name, person, firm, and role, for that person.
                                row = [last, person, role, firm, n]

                                # If something in each key varable then add to the dataset of people
                                if min(len(row[0]), len(row[1]), len(row[2]), len(row[3])) > 0:
                                    # Filter out people who already were added to the list
                                    if last in lsts:
                                        pass
                                    else:
                                        people.append(row)
                                        # remove duplicates
                                        people = [list(x) for x in set(tuple(x) for x in people)]
                                        # Append the sppl with this recent addition won't be added below
                                        # because the person's last name is not in the focal sentence,
                                        # which is the condition to be added for all other instances.
                                        sppl.append(people[-1])

                # Running list of full names of People list of lists
                fulls = [per[1] for per in people]
                lsts = [per[0] for per in people]

                # Create a list of people referenced in the focal sentence.
                if len(people) > 0:
                    for per in people:
                        last = per[0]
                        m = re.findall(r'\b' + last + r'\b', filtered_sent, re.IGNORECASE)
                        if len(m) > 0:
                            prow = [per[0], per[1], per[2], per[3], n]
                            prows.append(prow)
                            dfpeople.loc[len(dfpeople.index)] = [n, per[0], per[1], per[2], per[3]]
                            dfpeople = dfpeople.drop_duplicates(subset=['last', 'full'], keep='first')\
                                .reset_index(drop=True)


                # Cumulative list of people from people list referenced in sentences
                # Will refer to this in cases of pronouns (e.g., "she said"), etc.
                for per in people:
                    y = re.findall(r'\b' + per[0] + r'\b', filtered_sent, re.IGNORECASE)
                    if len(y) > 0:
                        per[4] = n
                        sppl.append(per)

                # Carry forward the prows column until replaced with next person
                # In essence, prows tells you the most recent person in each sentence including the focal one.
                if len(prows)==0 and len(dfsent.index)>0:
                        prows = dfsent.loc[len(dfsent.index) - 1].at['prows']

                covered = None



                dfsent.loc[len(dfsent.index)] = [n, stence, filtered_sent,sentiment, orgs, corgs, ppl, prows, people, covered]

                n += 1
                h.n += 1



            ############################################################################################
            # NOW YOU HAVE THE LIST, TAG SENTENCES WITH PEOPLE AND QUOTES
            ############################################################################################
            # Reset count variables
            n = 1
            h.n = 1
            for se in dfsent.index:

                # Assign the various lists and variables for this sentence from the dfsent dataframe
                stence = dfsent["stence"][se]
                h.stence = stence
                filtered_sent = dfsent["filtered_sent"][se]
                h.sentiment = dfsent["sentiment"][se]
                orgs = dfsent["orgs"][se]
                corgs = dfsent["orgs"][se]
                ppl = dfsent["ppl"][se]
                prows = dfsent["prows"][se]
                people = dfsent["people"][se]



                if dfsent["covered"][se] == 1:
                    info_extraction(stence, orgs)
                    n += 1
                    h.n += 1
                    continue


                # Deploy function to extract information from the sentence
                info_extraction(stence, orgs)

                most_recent_person = ""
                next_person = ""
                # Establish the last person referenced (first filter later ones and then take the last one)
                if len(prows)>0:
                    most_recent_person = prows[0]


                #############################
                elif len(prows)==0:
                    df_first_person = dfsent[dfsent['prows'].str.len()>0]
                    next_person = df_first_person["prows"][0]
                    print (n, df_first_person)
##############################





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

                        # Since no clear person reference, check for FULL quotes:

                        x1 = re.findall(r'(' + start + ')(.+)(' + end + ')', stence, re.IGNORECASE)
                        x1 = ''.join(''.join(elems) for elems in x1)
                        x1 = re.sub(r'(' + qs + r')', '', x1)

                        if len(x1) > 0:
                            # if there is a said equivalent and quote but no person, check if pronoun near said equivalent
                            k1 = re.findall(
                                r'(' + stext + r')\W+(?:\w+\W+){0,' + str(mxpron) + r'}?(' + pron + r')',
                                filtered_sent, re.IGNORECASE)
                            k2 = re.findall(
                                r'(' + pron + r')\W+(?:\w+\W+){0,' + str(mxpron) + r'}?(' + stext + r')',
                                filtered_sent, re.IGNORECASE)
                            k = k1 + k2
                            k = ''.join(''.join(elems) for elems in k)
                            k = re.sub(r'(' + qs + r')', '', k)
                            if len(k) > 0:
                                h.quote = x1
                                h.last = person_set(most_recent_person,next_person)[0]
                                h.last = person_set(most_recent_person,next_person)[1]
                                h.last = person_set(most_recent_person,next_person)[2]
                                h.last = person_set(most_recent_person,next_person)[3]
                                h.qtype = "single sentence quote"
                                h.qsaid = "yes"
                                h.qpref = "no"
                                h.comment = "pronoun"
                                h.todf()
                                dx += 1

                            # If said equiv and quote, but no people or pronoun, check if organization listed, as in "BMW's spokesman said..."
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
                                        # The person would be the representative's reference, removing the org name
                                        h.person = "NA"
                                        h.firm = org
                                        h.role = re.sub(r'' + org + '', '', k)
                                        h.quote = x1
                                        h.qtype = "single sentence quote"
                                        h.qsaid = "yes"
                                        h.qpref = "no"
                                        h.comment = "Org representative"
                                        h.todf()
                                        dx += 1


                                    elif len(k) == 0:
                                        h.last = "NA"
                                        h.person = "NA"
                                        h.firm = org
                                        h.role = "NA"
                                        h.quote = x1
                                        h.qtype = "single sentence quote"
                                        h.qsaid = "yes"
                                        h.qpref = "no"
                                        h.comment = "from org, nondiscriptives"
                                        h.todf()
                                        dx += 1
                                elif len(orgs) > 1:

                                    h.last = "NA"
                                    h.person = "NA"
                                    h.firm = ''.join(''.join(elems) for elems in orgs)
                                    h.role = "NA"
                                    h.quote = x1
                                    h.qtype = "Information - non-discript single sentence quote"
                                    h.qsaid = "yes"
                                    h.qpref = "no"
                                    h.comment = "Review - from nondiscriptives"
                                    h.todf()
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
                                        m = int(sppl[-1][-1])

                                        l = n - int(m)

                                        # If person in the previous sentence, use that one
                                        if l == 1:
                                            h.quote = x1
                                            h.last = person_set(most_recent_person,next_person)[0]
                                            h.last = person_set(most_recent_person,next_person)[1]
                                            h.last = person_set(most_recent_person,next_person)[2]
                                            h.last = person_set(most_recent_person,next_person)[3]

                                            h.qtype = "single sentence quote"
                                            h.qsaid = "yes"
                                            h.qpref = "no"
                                            h.comment = "pronoun"
                                            h.todf()
                                            dx += 1
                                        # if no person in previous sentence, then check if quote in previous sentence
                                        elif l > 1:
                                            # Check if there's something in DF:
                                            if dx>0:
                                                s = df.loc[len(df.index) - 1].at['sent']
                                                t = n - int(s)
                                                if t == 1:
                                                    h.quote = x1
                                                    h.last = df.loc[len(df.index) - 1].at['last']
                                                    h.person = df.loc[len(df.index) - 1].at['person']
                                                    h.firm = df.loc[len(df.index) - 1].at['firm']
                                                    h.role = df.loc[len(df.index) - 1].at['role']
                                                    h.qtype = "single sentence quote"
                                                    h.qsaid = "yes"
                                                    h.qpref = "no"
                                                    h.comment = "pronoun - reference partial profile"
                                                    h.todf()
                                                    dx += 1



                        # If said equivalent, no people and no quote, check if plural present
                        elif len(x1) == 0:
                            m = re.findall(r'(' + plurals + r')', filtered_sent, re.IGNORECASE)

                            if len(m) > 0:
                                h.last = "NA"
                                h.person = "NA"
                                h.firm = "NA"
                                h.role = ''.join(''.join(elems) for elems in m)
                                h.quote = stence
                                h.qtype = "paraphrase"
                                h.qsaid = "yes"
                                h.qpref = "no"
                                h.comment = "from plural individuals"
                                h.todf()
                                dx += 1


                            # If said equivalent, no people, no quote, no plural, check if pronoun, rep, and finally
                            # Org
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
                                if len(k) > 0:
                                    h.quote = stence
                                    h.last = person_set(most_recent_person,next_person)[0]
                                    h.last = person_set(most_recent_person,next_person)[1]
                                    h.last = person_set(most_recent_person,next_person)[2]
                                    h.last = person_set(most_recent_person,next_person)[3]

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
                                        if len(role_list) > 0:

                                            h.quote = stence
                                            h.last = person_set(most_recent_person,next_person)[0]
                                            h.last = person_set(most_recent_person,next_person)[1]
                                            h.last = person_set(most_recent_person,next_person)[2]
                                            h.last = person_set(most_recent_person,next_person)[3]

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
                                                            h.qtype = "paraphrase"
                                                            h.qsaid = "yes"
                                                            h.qpref = "no"
                                                            h.comment = "Review - paraphrase from multiple firms"
                                                            h.todf()
                                                            dx += 1





                        # Because said equivalent but no full quote, check if it's a beginning of
                        # multi-sentence quote
                        else:
                            # Run multi-sentence function and get outputs
                            last_sindex = multisent(stence)[0]
                            msent_quote = multisent(stence)[1]
                            last_after_quote = multisent(stence)[2]
                            first_said = multisent(stence)[3]
                            first_pron = multisent(stence)[4]
                            first_ppl = multisent(stence)[5]
                            last_said = multisent(stence)[6]
                            last_pron = multisent(stence)[7]
                            last_ppl = multisent(stence)[8]
                            msent_quotes = multisent(stence)[9]
                            stences = multisent(stence)[10]
                            sentiments = multisent(stence)[11]
                            nindex = multisent(stence)[12]
                            psent_index = n-2

                            # If there is a multi-sentence quote, then do stuff
                            if len(msent_quote)> 0:
                                # We know there is a said equivalent in the first sentence
                                # So now it's just checking who said it.
                                # First, check if person:
                                if len (first_ppl) > 0:
                                    if len(first_pron) == 0:
                                        for i in range(0, len(msent_quotes)-1):

                                            h.quote = msent_quotes[i]
                                            h.stence = stences[i]
                                            h.sentiment = sentiments[i]
                                            h.n = nindex[i]
                                            # Then it must be the most recent person, who is from this sentence
                                            h.last = person_set(most_recent_person,next_person)[0]
                                            h.last = person_set(most_recent_person,next_person)[1]
                                            h.last = person_set(most_recent_person,next_person)[2]
                                            h.last = person_set(most_recent_person,next_person)[3]
                                            h.qtype = "multi-sentence quote"
                                            h.qsaid = "yes"
                                            h.qpref = "no"
                                            h.comment = "Person named in first sentence"
                                            h.todf()
                                            h.n = n

                                    elif len(first_pron) > 0:
                                        for i in range(0, len(msent_quotes)):
                                            h.quote = msent_quotes[i]
                                            h.stence = stences[i]
                                            # Then it must be the most recent person
                                            h.last = person_set(most_recent_person,next_person)[0]
                                            h.last = person_set(most_recent_person,next_person)[1]
                                            h.last = person_set(most_recent_person,next_person)[2]
                                            h.last = person_set(most_recent_person,next_person)[3]
                                            h.qtype = "multi-sentence quote"
                                            h.qsaid = "yes"
                                            h.qpref = "no"
                                            h.comment = "Pronoun"
                                            h.todf()


                    # if said something and exactly one person from list, do stuff
                    elif len(ppl) == 1:
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

                            quote = s1
                            quote = re.sub(r'^' + stext + r'.+', '', quote)
                            h.quote = re.sub(r'' + stext + r'\Z', '', quote)
                            h.last = prows[0][0]
                            h.person = prows[0][1]
                            h.role = prows[0][2]
                            h.firm = prows[0][3]

                            h.qtype = "single sentence quote"
                            h.qsaid = "yes"
                            h.qpref = "single"
                            h.comment = "high accuracy"
                            h.todf()
                            dx += 1


                        # Check if the multiple sentence quote
                        elif len(s1) == 0:
                            # Because said equivalent but no full quote, check
                            # if there is a beginning of a multi-sentence quote. Must see if already covered
                            msent_quote = multisent(stence)[1]
                            msent_quotes = multisent(stence)[9]
                            stences = multisent(stence)[10]
                            sentiments = multisent(stence)[11]
                            nindex = multisent(stence)[12]

                            if len(msent_quote) > 0:
                                for i in range(0, len(msent_quotes)-1):
                                    h.quote = msent_quotes[i]
                                    h.stence = stences[i]
                                    h.sentiment = sentiments[i]
                                    h.n = nindex[i]
                                    h.last = prows[0][0]
                                    h.person = prows[0][1]
                                    h.role = prows[0][2]
                                    h.firm = prows[0][3]
                                    h.qtype = "multi-sentence quote"
                                    h.qsaid = "yes"
                                    h.qpref = "single"
                                    h.comment = "high accuracy"
                                    h.todf()
                                    h.n = n
                                    dx += 1
                            # If there is no full quote and no multi-sentence quote, but there is an ending quote and a person,
                            # update the previous row with this person. The previous row was a multi-sentence quote defaulted
                            # defaulted as the most previous person mentioned. So this will just confirm the previous person,
                            # or update it appropriately
                            elif len(msent_quote) == 0:
                                q1a = re.findall(r'(' + end + r').+(' + stext + r').+(' + last + r')', stence,
                                                 re.IGNORECASE)
                                q1b = re.findall(r'(' + end + r').+(' + last + r').+(' + stext + r')', stence,
                                                 re.IGNORECASE)
                                q1 = q1a + q1b
                                q1 = ''.join(''.join(elems) for elems in q1)
                                q = re.sub(r'(' + qs + r')', '', q1)

                                if len(q) > 0:
                                    df.loc[len(df.index) - 1].at['last'] = prows[0][0]
                                    df.loc[len(df.index) - 1].at['person'] = prows[0][1]
                                    df.loc[len(df.index) - 1].at['firm'] = prows[0][2]
                                    df.loc[len(df.index) - 1].at['role'] = prows[0][3]

                                elif len(q1) == 0:
                                    h.quote = stence
                                    h.last = prows[0][0]
                                    h.person = prows[0][1]
                                    h.role = prows[0][2]
                                    h.firm = prows[0][3]
                                    h.qtype = "paraphrase"
                                    h.qsaid = "yes"
                                    h.qpref = "single"
                                    h.comment = "high accuracy"
                                    h.todf()
                                    dx += 1


                    # If said equivalent but more than one person, do stuff:
                    elif len(ppl) > 1:
                        ##################################
                        # MORE THAN ONE PEOPLE REFERENCED
                        ##################################
                        for per in prows:
                            h.last = per[0]
                            h.person = per[1]
                            h.role = per[2]
                            h.firm = per[3]

                            # Check if full quote
                            # Quote then said-equivalent is within three spaces of last name
                            s1a = re.findall(
                                r'(' + start + r')(.+)(' + end + r')(?=.+(' + h.last + r')\W+(?:\w+\W+){0,' + str(
                                    mxlds) + r'}?' + stext + r')', stence, re.IGNORECASE)
                            # Said equiv then last name then a quote (distance not matter here because sometimes full title in between)
                            s1b = re.findall(
                                r'(' + h.last + r').+(' + stext + 'r).+(' + start + r')(.+)(' + end + r')',
                                stence,
                                re.IGNORECASE)
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
                                if len(q1) > 0:
                                    for s in range(n, len(doc.sentences)):
                                        j = doc.sentences[s]
                                        t = j.text
                                        q2 = re.findall(r'(.+)(' + end + r')', t, re.IGNORECASE)
                                        q2 = ''.join(''.join(elems) for elems in q2)
                                        q2 = re.sub(r'(' + qs + r')', '', q2)

                                        if len(q2) == 0:
                                            q1 = q1 + " " + t
                                        if len(q2) > 0:
                                            q1 = q1 + " " + q2
                                            break
                                    quote = q1
                                    quote = re.sub(r'^' + stext + r'.+', '', quote)
                                    h.quote = re.sub(r'' + stext + r'\Z', '', quote)
                                    h.qtype = "multi-sentence quote"
                                    h.qsaid = "yes"
                                    h.qpref = "multiple"
                                    h.comment = "review"
                                    h.todf()
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

                        if len(x1) > 0:
                            quote = x1
                            quote = re.sub(r'^' + stext + r'.+', '', quote)
                            h.quote = re.sub(r'' + stext + r'\Z', '', quote)
                            h.last = person_set(most_recent_person,next_person)[0]
                            h.last = person_set(most_recent_person,next_person)[1]
                            h.last = person_set(most_recent_person,next_person)[2]
                            h.last = person_set(most_recent_person,next_person)[3]

                            h.qtype = "single sentence quote"
                            h.qsaid = "no"
                            h.qpref = "no"
                            h.comment = "Quote with at least two words inside"
                            h.todf()
                            dx += 1


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
                                if len(multisent(stence)[1])>0:
                                    last_sindex = multisent(stence)[0]
                                    msent_quote = multisent(stence)[1]
                                    last_after_quote = multisent(stence)[2]
                                    first_said = multisent(stence)[3]
                                    first_pron = multisent(stence)[4]
                                    first_ppl = multisent(stence)[5]
                                    last_said = multisent(stence)[6]
                                    last_pron = multisent(stence)[7]
                                    last_ppl = multisent(stence)[8]
                                    msent_quotes = multisent(stence)[9]
                                    stences = multisent(stence)[10]
                                    sentiments = multisent(stence)[11]
                                    nindex= multisent(stence)[12]
                                    psent_index = n - 2

                                    if len(msent_quote) > 0:
                                        # Check if there's a person in that last sentence
                                        if len(last_ppl) > 0:
                                            for i in range(0, len(msent_quotes)):
                                                h.quote = msent_quotes[i]
                                                h.stence = stences[i]
                                                h.sentiment = sentiments[i]
                                                h.n = nindex[i]
                                                h.last = dfsent.loc[last_sindex].at['people'][0][0]
                                                h.person = dfsent.loc[last_sindex].at['people'][0][1]
                                                h.role = dfsent.loc[last_sindex].at['people'][0][2]
                                                h.firm = dfsent.loc[last_sindex].at['people'][0][3]
                                                h.qtype = "multi-sentence quote"
                                                h.qsaid = "no"
                                                h.qpref = "no"
                                                h.comment = "high accuracy"
                                                h.todf()
                                                h.n = n
                                                dx += 1

                                        elif len(last_ppl) == 0:
                                            for i in range(0, len(msent_quotes)):
                                                h.quote = msent_quotes[i]
                                                h.stence = stences[i]
                                                h.sentiment = sentiments[i]
                                                h.n = nindex[i]
                                                h.last = person_set(most_recent_person,next_person)[0]
                                                h.last = person_set(most_recent_person,next_person)[1]
                                                h.last = person_set(most_recent_person,next_person)[2]
                                                h.last = person_set(most_recent_person,next_person)[3]
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
                    elif len(ppl) > 0:
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
                                quote = x1
                                quote = re.sub(r'^' + stext + r'.+', '', quote)
                                h.quote = re.sub(r'' + stext + r'\Z', '', quote)
                                h.last = "NA"
                                h.person = "NA"
                                h.firm = "NA"
                                h.role = "NA"
                                h.qtype = "single sentence quote"
                                h.qsaid = "no"
                                h.qpref = "one or more"
                                h.comment = "Review, whether quote or something else, not clear"
                                h.todf()
                                dx += 1
                        # Because said equivalent but no full quote, check if it's a beginning of multi-sentence quote
                        # If so, assume it would have to be from the most recent referenced person
                        # As in "he said", referring to the most recent person
                        elif len(x1) == 0:
                            # check if a quote begins in this sentence:
                            q1 = re.findall(r'(' + start + r')(.+)', stence, re.IGNORECASE)
                            q1 = ''.join(''.join(elems) for elems in q1)
                            q1 = re.sub(r'(' + qs + r')', '', q1)

                            # If a quote does begin in this sentence, concatinate the subsequent sentences until the quote ends
                            if len(q1) > 0:
                                for s in range(n, len(doc.sentences)):
                                    j = doc.sentences[s]
                                    t = j.text
                                    q2 = re.findall(r'(.+)(' + end + r')', t, re.IGNORECASE)
                                    q2 = ''.join(''.join(elems) for elems in q2)
                                    q2 = re.sub(r'(' + qs + r')', '', q2)
                                    if len(q2) == 0:
                                        q1 = q1 + " " + t
                                    if len(q2) > 0:
                                        q1 = q1 + " " + q2
                                        break
                                quote = q1
                                quote = re.sub(r'^' + stext + r'.+', '', quote)
                                h.quote = re.sub(r'' + stext + r'\Z', '', quote)
                                h.last = person_set(most_recent_person,next_person)[0]
                                h.last = person_set(most_recent_person,next_person)[1]
                                h.last = person_set(most_recent_person,next_person)[2]
                                h.last = person_set(most_recent_person,next_person)[3]
                                h.qtype = "multi-sentence quote"
                                h.qsaid = "no"
                                h.qpref = "one or more"
                                h.comment = "high accuracy - prior person referenced"
                                h.todf()
                                dx += 1
                df
                n += 1
                h.n += 1



        #except:
        #   df.loc[len(df.index)] = [n, "Issue", "Issue", "Issue", "Issue", "Issue", "Issue", "Issue", "Issue",
        #                            "Sentence-level Issue", AN, date, source, regions, subjects, misc, industries,
        #                            focfirm, by, comp, sentiment, stence]

        t2 = time.time()
        total = t1 - t0
        print(art + 1, h.AN, "time run:", round(t2 - t1, 2), "total hours:", round((t2 - t0) / (60 * 60), 2),
              "mean rate:", round((t2 - t0) / (art + 1), 2))

    #except:
    #   df.loc[len(df.index)] = [n, "Issue", "Issue", "Issue", "Issue", "Issue", "Issue", "Issue", "Issue",
    #                            "Article NLP issue", AN, date, source, "Issue", "Issue", "Issue", "Issue", "Issue",
    #                            "Issue", "Issue", "Issue", "Issue"]




# output the chunk
# OUTFILE = f"{OUTDIR}/{INFILE.split('.')[0]}_df_quotes_{STARTROW.zfill(6)}_{ENDROW.zfill(6)}.xlsx"
# df.to_excel(OUTFILE)


df

# print(people)


# TO DO:
# "Former" leaders
# Multiple quotes in a sentence
# Test 25 random ones.
