import time
from tkinter import *
import datetime
import os
import glob
import timeit
import striprtf
import PyRTF
import glob
import re
import sqlite3
import csv
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np
import _pickle
import time
import stanza

# stanza.download('en')
nlp = stanza.Pipeline('en')

dfa = pd.read_excel(r'C:\Users\danwilde\Dropbox (Penn)\Dissertation\Factiva\final4.xlsx')

# combine the lead paragraph and the body of the article
i = 5

text = dfa['LP'][i] + " " + dfa['TD'][i]
firm = dfa["Firm"][i]
vertexid = dfa["vertex.id"][i]
start = dfa["start"][i]
end = dfa["end"][i]
AN = dfa["AN"][i]
date = dfa["PD"][i]
source = dfa["SN"][i]
regions = dfa["RE"][i]
subjects = dfa["NS"][i]
misc = dfa["IPD"][i]
industries = dfa["IN"][i]

columns = ['sent', 'person', 'last', 'role', 'firm', 'quote', 'qtype', 'qsaid', 'qpref', 'comment', 'AN', 'date',
           'source','sentence']
df = pd.DataFrame(columns=columns)


stext = r'said|n.t say|say|told|n.t tell\S+|tell\S+|assert\S+|according\s+to|n.t comment\S+|comment\S+|quote\S+' \
        r'|describ\S+|n.t communicat\S+|communicat\S+|articulat\S+|n.t divulg\S+|divulg\S+|noted|noting|espressed' \
        r'|recounted|suggested|explained|added'

l1 = r'vice chair\S+|chair\S+|Governor\b|Congress\S+|\S+\s+head\s+of\s\S+|chair\S+\s+and\s+\S+\s+head\b|banker\b' \
     r'|specialist\b|accountant\b|journalist\b|reporter\b|analyst\b|consultant\b|boss\b' \
     r'|executive\s+vice\s+president\s+of\s\S+|executive\s+vice\s+president\b|vice\s+president\b|president\b' \
     r'|general\s+manager\b|senior\s+manager\b' \
     r'|sr\s+manager\b|manager\b|vp\s+of\s+\S+|vp\b|director\s+of\s+\S+|director\b|chief.+officer' \
     r'| president\s+and\s+\S+|\S+\s+executive\b\s+of\s+the\s+\S+|chief\s+executive\b\s+of\s+|chief\s+executive\b'


l2 = r'C.O\b'
pron = r'he|she'
plurals = r'analysts|bankers|consultants|executives|management|participants'
rep = r'spokesman\b|spokesperson\b|spokeswoman\b|representative\b|official\b|executive\b'
start = r'\"|\“|\`\`'
end = r'\"|\”|\'\''
qs = r'\“|\`\`|\”|\"|\'\''

d = re.findall(r'\(.+\)\s*(--|-)', text, re.IGNORECASE)
e = re.findall(r'^.+/\b.+\b/\s*(--|-)', text, re.IGNORECASE)

if len(d) > 0:
    text = re.sub(r'^.+\(.+\)\s*(--|-)', '', text)

elif len(e) > 0:
    text =  re.sub(r'^.+/\b.+\b/\s*-', '',text)

doc = nlp(text)

people = []
sppl = []
fulls = []
lsts = []



# Set the max number of words away from the focal leader and the focal organization
mxld = 6
mxorg = 6
mxlds = 2
mxpron = 2

n = 1

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

    stence = sent.text

    # Check if you can add to dataset of people/firm/role
    # First, compile a list of people and organizations within the sentence.
    for ent in sent.ents:
        if ent.type == "PERSON":
            ppl.append(ent.text)
            l = ent.text.split()[-1]
            lasts.append(l)
        if ent.type == "ORG":
            orgs.append(ent.text)

            # Next, if there are any people, see if should add to running People list
    if len(ppl) > 0:
        # see if the person's title is mentioned
        for per in ppl:
            # reset key variables each person reference
            last = ""
            person = ""
            firm = ""
            role = ""

            pos1a = re.findall(r'\b(?:' + per + r'\W+(?:\w+\W+){0,' + str(mxld) + r'}?(' + l1 + r'))', stence,
                               re.IGNORECASE)
            pos1b = re.findall(r'\b(?:' + per + r'\W+(?:\w+\W+){0,' + str(mxld) + r'}?(' + l2 + r'))', stence)
            pos2a = re.findall(r'' + l1 + r'\W+(?:\w+\W+){0,' + str(mxld) + r'}?' + per + r'\b', stence)
            pos2b = re.findall(r'' + l2 + r'\W+(?:\w+\W+){0,' + str(mxld) + r'}?' + per + r'\b', stence)
            pos = pos1a + pos1b + pos2a + pos2b





            # If the title is mentioned, move to next step

            if len(pos) > 0:
                # If there is an organization mentioned, check if near to reference of person
                if len(orgs) > 0:
                    for org in orgs:

                        org = org.replace(r"'s", "")
                        org1 = re.findall(r'\b(?:' + per + r'\W+(?:\w+\W+){0,' + str(mxld) + r'}?(' + org + r'))',
                                          stence, re.IGNORECASE)
                        org2 = re.findall(r'(' + org + r')\W+(?:\w+\W+){0,' + str(mxld) + r'}?' + per + r'\b', stence,
                                          re.IGNORECASE)
                        f = org1 + org2
                        f = ''.join(''.join(elems) for elems in f)

                        # If yes,
                        if len(f) > 0:
                            person = per
                            last = person.split()[-1]
                            r = pos[0]
                            r = re.sub(r'(' + per + r')', '', r)
                            role = re.sub(r',', '', r)
                            firm = f
                #If there's a titled person with no org, it's may be the org in previous reference
                # If a repeat person, this will be filtered out below.
                elif len(orgs)==0:
                    person = per
                    last = person.split()[-1]
                    r = pos[0]
                    r = re.sub(r'(' + per + r')', '', r)
                    role = re.sub(r',', '', r)
                    firm = df.loc[len(df.index) - 1].at['firm']

            # Combine last name, person, firm, and role, for that person.
            row = [last, person, firm, role]

            # If something in each key varable then add to the dataset of people
            if min(len(row[0]), len(row[1]), len(row[2]), len(row[3])) > 0:
                #Filter out people who already were added to the list
                if last in lsts:
                    pass
                else:
                    people.append(row)
                    # remove duplicates
                    people = [list(x) for x in set(tuple(x) for x in people)]

    # Running list of full names of People list of lists
    fulls = [per[1] for per in people]
    lsts = [per[0] for per in people]

    # Create a list of people referenced in the focal sentence.
    if len(people) > 0:
        for per in people:
            last = per[0]
            m = re.findall(r'\b' + last + r'\b', stence, re.IGNORECASE)
            if len(m) > 0:
                prow = [per[0], per[1], per[2], per[3]]
                prows.append(prow)

    # Cumulative list of people from people list referenced in sentences
    # Will refer to this in cases of pronouns (e.g., "she said"), etc.
    for per in people:
        y = re.findall(r'\b' + per[0] + r'\b', stence, re.IGNORECASE)
        if len(y) > 0:
            sppl.append(per)

    ############################################################################################
    # NOW YOU HAVE UPDATED YOUR RUNNING LISTS, TAG SENTENCES WITH PEOPLE AND QUOTES, IF APPLICABLE
    ############################################################################################
    ##################################
    # AT LEAST ONE PERSON IN CUMULATIVE LIST
    ##################################
    if len(people) > 0:
        ##################################
        # SENTENCES WITH A "SAID" EQUIVALENT
        ##################################
        s = re.findall(r'' + stext + '', stence, re.IGNORECASE)
        # First check if "said" equivalent in sentence
        if len(s) > 0:
            if len(prows) == 0:
                ##################################
                # NO PEOPLE REFERENCE
                ##################################
                if len(ppl) == 0:

                    # Since no clear person reference, check for FULL quotes:

                    x1 = re.findall(r'(' + start + ')(.+)(' + end + ')', stence, re.IGNORECASE)
                    x1 = ''.join(''.join(elems) for elems in x1)
                    x1 = re.sub(r'[' + qs + r']', '', x1)

                    if len(x1) > 0:
                        # if there is a said equivalent and quote but no person, check if pronoun near said equivalent
                        k1 = re.findall(r'(' + stext + r')\W+(?:\w+\W+){0,' + str(mxpron) + r'}?(' + pron + r')',
                                        stence, re.IGNORECASE)
                        k2 = re.findall(r'(' + pron + r')\W+(?:\w+\W+){0,' + str(mxpron) + r'}?(' + stext + r')',
                                        stence, re.IGNORECASE)
                        k = k1 + k2
                        k = ''.join(''.join(elems) for elems in k)
                        k = re.sub(r'[' + qs + r']', '', k)
                        if len(k) > 0:

                            quote = x1
                            quote = re.sub(r'^' + stext + '.+', '', quote)
                            quote = re.sub(r'' + stext + '\Z', '', quote)
                            last = sppl[-1][0]
                            person = sppl[-1][1]
                            firm = sppl[-1][2]
                            role = sppl[-1][3]
                            qtype = "single sentence quote"
                            qsaid = "yes"
                            qpref = "no"
                            comment = "pronoun"
                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment,
                                                     AN, date, source, stence]
                        # If said equiv and quote, but no people or pronoun, check if organization listed, as in "BMW's spokesman said..."
                        elif len(k) == 0:
                            if len(orgs) == 1:
                                # check if spokeperson
                                org = orgs[0]
                                k3 = re.findall(r'' + org + r'.+(' + rep + r')', stence, re.IGNORECASE)
                                k4 = re.findall(r'' + rep + r'.+(' + org + r')', stence, re.IGNORECASE)
                                k = k3 + k4
                                k = ''.join(''.join(elems) for elems in k)
                                k = re.sub(r'[' + qs + r']', '', k)
                                if len(k) > 0:
                                    last = "NA"
                                    # The person would be the representative's reference, removing the org name
                                    person = re.sub(r'' + org + '', '', k)
                                    firm = org
                                    role = "NA"
                                    quote = x1
                                    quote = re.sub(r'^' + stext + '.+', '', quote)
                                    quote = re.sub(r'' + stext + '\Z', '', quote)
                                    qtype = "single sentence quote"
                                    qsaid = "yes"
                                    qpref = "no"
                                    comment = "Org representative"
                                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref,
                                                             comment, AN, date, source, stence]
                                elif len(k) == 0:
                                    last = "NA"
                                    person = "NA"
                                    firm = orgs[0]
                                    role = "NA"
                                    quote = x1
                                    quote = re.sub(r'^' + stext + '.+', '', quote)
                                    quote = re.sub(r'' + stext + '\Z', '', quote)
                                    qtype = "single sentence quote"
                                    qsaid = "yes"
                                    qpref = "no"
                                    comment = "from org, nondiscriptives"
                                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref,
                                                             comment, AN, date, source, stence]
                            elif len(orgs) > 1:
                                last = "NA"
                                person = "NA"
                                firm = ''.join(''.join(elems) for elems in orgs)
                                role = "NA"
                                quote = x1
                                quote = re.sub(r'^' + stext + '.+', '', quote)
                                quote = re.sub(r'' + stext + '\Z', '', quote)
                                qtype = "Information - non-discript single sentence quote"
                                qsaid = "yes"
                                qpref = "no"
                                comment = "Review - from nondiscriptives"

                    # If said equivalent, no people and no quote, check if plural present
                    elif len(x1) == 0:
                        m = re.findall(r'(' + plurals + r')', stence, re.IGNORECASE)
                        if len(m) > 0:
                            last = "NA"
                            person = ' '.join(m)
                            firm = "NA"
                            role = "NA"
                            quote = stence
                            qtype = "paraphrase"
                            qsaid = "yes"
                            qpref = "no"
                            comment = "from plural individuals"
                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment,
                                                     AN, date, source, stence]

                        # If said equivalent, no people no quote and no plural, must be nondiscript said equivalent
                        elif len(m) == 0:
                            k1 = re.findall(r'(' + stext + r')\W+(?:\w+\W+){0,' + str(mxpron) + r'}?(' + pron + r')',
                                            stence, re.IGNORECASE)
                            k2 = re.findall(r'(' + pron + r')\W+(?:\w+\W+){0,' + str(mxpron) + r'}?(' + stext + r')',
                                            stence, re.IGNORECASE)
                            k = k1 + k2
                            k = ''.join(''.join(elems) for elems in k)
                            k = re.sub(r'[' + qs + r']', '', k)
                            if len(k) > 0:
                                quote = stence
                                # quote = re.sub(r'^'+stext+'.+','',quote)
                                # quote = re.sub(r''+stext+'\Z','',quote)
                                last = sppl[-1][0]
                                person = sppl[-1][1]
                                firm = sppl[-1][2]
                                role = sppl[-1][3]
                                qtype = "paraphrase"
                                qsaid = "yes"
                                qpref = "no"
                                comment = "pronoun"
                                df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref,
                                                         comment, AN, date, source, stence]
                            # If said equiv, but no person, no quote, not plural, and no pronoun, check if a title
                            # such as "The executive said..." or "the chairman said..."
                            elif len(k) == 0:
                                p1a = re.findall(r'(' + l1 + r')', stence, re.IGNORECASE)
                                p1b = re.findall(r'(' + l2 + r')', stence)
                                p1 = p1a + p1b
                                p1 = ''.join(''.join(elems) for elems in p1)
                                if len(p1)>0:
                                    quote = stence
                                    last = sppl[-1][0]
                                    person = sppl[-1][1]
                                    firm = sppl[-1][2]
                                    role = sppl[-1][3]
                                    qtype = "paraphrase"
                                    qsaid = "yes"
                                    qpref = "no"
                                    comment = "title reference"
                                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref,
                                                             comment, AN, date, source, stence]
                                elif len(p1)==0:

                                    last = "misc"
                                    person = "misc"
                                    firm = "misc"
                                    role = "misc"
                                    quote = stence
                                    qtype = "paraphrase"
                                    qsaid = "yes"
                                    qpref = "no"
                                    comment = "Review - paraphrase from nondiscriptives"
                                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref,
                                                             comment, AN, date, source, stence]



                    # Because said equivalent but no full quote, check if it's a beginning of multi-sentence quote
                    # If so, assume it would have to be from the most recent referenced person
                    # As in "he said", referring to the most recent person
                    else:
                        # check if a quote begins in this sentence:
                        q1 = re.findall(r'(' + start + r').+', stence, re.IGNORECASE)
                        q1 = ' '.join(q1)
                        q1 = re.sub(r'[' + qs + r']', '', q1)

                        # If a quote does begin in this sentence, concatinate the subsequent sentences until the quote ends
                        if len(q1) > 0:
                            for s in range(n, len(doc.sentences)):
                                j = doc.sentences[s]
                                t = j.text
                                q2 = re.findall(r'(.+)[' + end + r']', t, re.IGNORECASE)
                                q2 = ' '.join(q2)
                                q2 = re.sub(r'[' + qs + r']', '', q2)
                                if len(q2) == 0:
                                    q1 = q1 + " " + t
                                if len(q2) > 0:
                                    q1 = q1 + " " + q2
                                    break
                            quote = q1
                            quote = re.sub(r'^' + stext + '.+', '', quote)
                            quote = re.sub(r'' + stext + '\Z', '', quote)
                            last = sppl[-1][0]
                            person = sppl[-1][1]
                            firm = sppl[-1][2]
                            role = sppl[-1][3]
                            qtype = "multi-sentence quote"
                            qsaid = "yes"
                            qpref = "no"
                            comment = "high accuracy"
                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment,
                                                     AN, date, source, stence]


            # if said something and exactly one person from list, do stuff
            elif len(prows) == 1:
                ##################################
                # ONE AND ONLY ONE PEOPLE REFERENCE
                ##################################

                last = prows[0][0]

                # Check for FULL quotes:
                # first check for the FULL quote after the said equivalent
                # last name then said equivalent then quote
                s1a = re.findall(r'(' + last + r').+(' + stext + r').+(' + start + r')(.+)(' + end + r')', stence,
                                 re.IGNORECASE)
                # said equiv then last name then a quote
                s1b = re.findall(r'(' + stext + r').+(' + last + r').+(' + start + r')(.+)(' + end + r')', stence,
                                 re.IGNORECASE)
                # quote then said-equivalent then last name
                s2a = re.findall(r'(' + start + r')(.+)(' + end + r').+(' + stext + r').+(' + last + r')', stence,
                                 re.IGNORECASE)
                # quote then last name then said-equivalent
                s2b = re.findall(r'(' + start + r')(.+)(' + end + r').+(' + last + r').+(' + stext + r')', stence,
                                 re.IGNORECASE)
                s1 = s1a + s1b + s2a + s2b
                s1 = ''.join(''.join(elems) for elems in s1)
                s1 = re.sub(r'[' + qs + r']', '', s1)
                s1 = re.sub(r'(' + stext + r')', '', s1)
                s1 = re.sub(r'' + last + r'', '', s1)

                if len(s1) > 0:

                    quote = s1
                    quote = re.sub(r'^' + stext + r'.+', '', quote)
                    quote = re.sub(r'' + stext + r'\Z', '', quote)
                    last = prows[0][0]
                    person = prows[0][1]
                    firm = prows[0][2]
                    role = prows[0][3]
                    qtype = "single sentence quote"
                    qsaid = "yes"
                    qpref = "single"
                    comment = "high accuracy"
                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment, AN, date,
                                             source, stence]


                # Check if the multiple sentence quote
                elif len(s1) == 0:
                    # Because said equivalent but no full quote, check if it's a beginning of multi-sentence quote
                    # If so, assume it would have to be from the most recent referenced person
                    # check if a quote begins in this sentence:
                    q1 = re.findall(r'(' + last + r').+(' + stext + r').+(' + start + r')(.+)', stence, re.IGNORECASE)
                    q1 = ''.join(''.join(elems) for elems in q1)
                    q1 = re.sub(r'[' + qs + r']', '', q1)
                    q1 = re.sub(r'' + last + r'', '', q1)
                    q1 = re.sub(r'(' + stext + r')', '', q1)

                    # If a quote does begin in this sentence, concatinate the subsequent sentences until the quote ends
                    if len(q1) > 0:
                        for s in range(n, len(doc.sentences)):
                            j = doc.sentences[s]
                            t = j.text
                            q2 = re.findall(r'(.+)[' + end + r']', t, re.IGNORECASE)
                            q2 = ' '.join(q2)
                            q2 = re.sub(r'[' + qs + r']', '', q2)
                            if len(q2) == 0:
                                q1 = q1 + " " + t
                            if len(q2) > 0:
                                q1 = q1 + " " + q2
                                break
                        quote = q1
                        last = prows[0][0]
                        person = prows[0][1]
                        firm = prows[0][2]
                        role = prows[0][3]
                        qtype = "multi-sentence quote"
                        qsaid = "yes"
                        qpref = "single"
                        comment = "high accuracy"
                        df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment, AN,
                                                 date, source, stence]
                    # Check if the name is actually inside the quote
                    # If so, the speaker should be the previous referenced person
                    elif len(q1) == 0:
                        s3 = re.findall(r'(' + start + r')(.+' + last + r'.+)(' + end + r')', stence, re.IGNORECASE)
                        s3 = ''.join(''.join(elems) for elems in s3)
                        s3 = re.sub(r'[' + qs + r']', '', s3)

                        if len(s3) > 0:
                            last = sppl[-1][0]
                            person = sppl[-1][1]
                            firm = sppl[-1][2]
                            role = sppl[-1][3]
                            quote = s3
                            quote = re.sub(r'^' + stext + r'.+', '', quote)
                            quote = re.sub(r'' + stext + r'\Z', '', quote)
                            qtype = "single sentence quote"
                            qsaid = "yes"
                            qpref = "single"
                            comment = "Other person within the quote"
                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment,
                                                     AN, date, source, stence]

                        elif len(s3) == 0:
                            q1a = re.findall(r'(' + end + r').+(' + stext + r').+(' + last + r')', stence,
                                             re.IGNORECASE)
                            q1b = re.findall(r'(' + end + r').+(' + last + r').+(' + stext + r')', stence,
                                             re.IGNORECASE)
                            q1 = q1a + q1b
                            q1 = ''.join(''.join(elems) for elems in q1)
                            q = re.sub(r'[' + qs + r']', '', q1)
                            # If there is no full quote and no beginning quote, but there is an ending quote and a person,
                            # update the previous row with this person. The previous row was a multi-sentence quote defaulted
                            # defaulted as the most previous person mentioned. So this will just confirm the previous person,
                            # or update it appropriately
                            if len(q) > 0:
                                df.loc[len(df.index) - 1].at['last'] = prows[0][0]
                                df.loc[len(df.index) - 1].at['person'] = prows[0][1]
                                df.loc[len(df.index) - 1].at['firm'] = prows[0][2]
                                df.loc[len(df.index) - 1].at['role'] = prows[0][3]

                            elif len(q1) == 0:
                                quote = stence
                                last = prows[0][0]
                                person = prows[0][1]
                                firm = prows[0][2]
                                role = prows[0][3]
                                qtype = "paraphrase"
                                qsaid = "yes"
                                qpref = "single"
                                comment = "high accuracy"
                                df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref,
                                                         comment, AN, date, source, stence]


            # If said equivalent but more than one person, do stuff:
            elif len(prows) > 1:
                ##################################
                # MORE THAN ONE PEOPLE REFERENCED
                ##################################
                for per in prows:
                    last = per[0]
                    person = per[1]
                    firm = per[2]
                    role = per[3]
                    # Check if full quote
                    # Quote then said-equivalent is within three spaces of last name
                    s1a = re.findall(r'(' + start + r')(.+)(' + end + r')(?=.+(' + last + r')\W+(?:\w+\W+){0,' + str(
                        mxlds) + r'}?' + stext + r')', stence, re.IGNORECASE)
                    # Said equiv then last name then a quote (distance not matter here because sometimes full title in between)
                    s1b = re.findall(r'(' + last + r').+(' + stext + 'r).+(' + start + r')(.+)(' + end + r')', stence,
                                     re.IGNORECASE)
                    s1 = s1a + s1b
                    s1 = ''.join(''.join(elems) for elems in s1)
                    s1 = re.sub(r'[' + qs + r']', '', s1)
                    s1 = re.sub(r'(' + stext + r')', '', s1)
                    s1 = re.sub(r'' + last + r'', '', s1)
                    if len(s1) > 0:
                        quote = s1
                        quote = re.sub(r'^' + stext + r'.+', '', quote)
                        quote = re.sub(r'' + stext + r'\Z', '', quote)
                        qtype = "single sentence quote"
                        qsaid = "yes"
                        qpref = "multiple"
                        comment = "review"
                        df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment, AN,
                                                 date, source, stence]
                    # check if a quote begins in this sentence:
                    elif len(s1) == 0:
                        q1a = re.findall(r'(' + last + r').+(' + stext + r').+(' + start + r')(.+)', stence,
                                         re.IGNORECASE)
                        # said equiv then last name then a quote
                        q1b = re.findall(r'(' + stext + r').+(' + last + r').+(' + start + r')(.+)', stence,
                                         re.IGNORECASE)
                        q1 = q1a + q1b
                        q1 = ''.join(''.join(elems) for elems in q1)
                        q1 = re.sub(r'[' + qs + r']', '', q1)
                        q1 = re.sub(r'(' + stext + r')', '', q1)
                        q1 = re.sub(r'' + last + r'', '', q1)

                        # If a quote does begin in this sentence, concatinate the subsequent sentences until the quote ends
                        if len(q1) > 0:
                            for s in range(n, len(doc.sentences)):
                                j = doc.sentences[s]
                                t = j.text
                                q2 = re.findall(r'(.+)[' + end + r']', t, re.IGNORECASE)
                                q2 = ' '.join(q2)
                                q2 = re.sub(r'[' + qs + r']', '', q2)

                                if len(q2) == 0:
                                    q1 = q1 + " " + t
                                if len(q2) > 0:
                                    q1 = q1 + " " + q2
                                    break
                            quote = q1
                            quote = re.sub(r'^' + stext + r'.+', '', quote)
                            quote = re.sub(r'' + stext + r'\Z', '', quote)
                            qtype = "multi-sentence quote"
                            qsaid = "yes"
                            qpref = "multiple"
                            comment = "review"
                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment,
                                                     AN, date, source, stence]
                        elif len(q1) == 0:
                            quote = stence
                            qtype = "paraphrase"
                            qsaid = "yes"
                            qpref = "multiple"
                            comment = "review"
                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment,
                                                     AN, date, source, stence]


        elif len(s) == 0:
        ##################################
        # SENTENCES WITH NO "SAID" EQUIVALENT
        ##################################
            # Check if no person:
            if len(prows) == 0:
                ##################################
                # NO PEOPLE REFERENCE
                ##################################
                # Since no clear person reference, check for FULL quotes:
                x1a = re.findall(r'\"(\b.+\b\s+\b.+\b).+\"', stence, re.IGNORECASE)
                x1b = re.findall(r'``(\b.+\b\s+\b.+\b).+\"', stence, re.IGNORECASE)
                x1c = re.findall(r'``(\b.+\b\s+\b.+\b).+\'\'', stence, re.IGNORECASE)
                x1 = x1a + x1b + x1c
                x1 = ''.join(''.join(elems) for elems in x1)

                # if so, assume must be from the most recent person referenced in the article
                # As in "he said", refering to the most recent person
                if len(x1) > 0:
                    quote = x1
                    quote = re.sub(r'^' + stext + r'.+', '', quote)
                    quote = re.sub(r'' + stext + r'\Z', '', quote)
                    last = sppl[-1][0]
                    person = sppl[-1][1]
                    firm = sppl[-1][2]
                    role = sppl[-1][3]
                    qtype = "single sentence quote"
                    qsaid = "no"
                    qpref = "no"
                    comment = "Quote with at least two words inside"
                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment, AN, date,
                                             source, stence]
                # Because said equivalent but no full quote, check if it's a beginning of multi-sentence quote
                # If so, assume it would have to be from the most recent referenced person
                # As in "he said", referring to the most recent person
                elif len(x1) == 0:

                    # check if a quote begins in this sentence:
                    q1 = re.findall(r'(' + start + r')(.+)', stence, re.IGNORECASE)
                    q1 = ''.join(''.join(elems) for elems in q1)
                    q1 = re.sub(r'[' + qs + r']', '', q1)

                    # If a quote does begin in this sentence, concatinate the subsequent sentences until the quote ends
                    if len(q1) > 0:
                        for s in range(n, len(doc.sentences)):
                            j = doc.sentences[s]
                            t = j.text
                            q2 = re.findall(r'(.+)[' + end + r']', t, re.IGNORECASE)
                            q2 = ''.join(''.join(elems) for elems in q2)
                            q2a = re.findall(r'(.+)[' + end + r'](.+)', t, re.IGNORECASE)
                            q2a = ''.join(''.join(elems) for elems in q2a)
                            if len(q2) == 0:
                                q1 = q1 + " " + t
                            if len(q2) > 0:
                                q1 = q1 + " " + q2
                                q1a = q1 + " " + q2a
                                break

                        quote = q1
                        last = sppl[-1][0]
                        person = sppl[-1][1]
                        firm = sppl[-1][2]
                        role = sppl[-1][3]
                        qtype = "multi-sentence quote"
                        qsaid = "no"
                        qpref = "no"
                        comment = "high accuracy"
                        df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment, AN,
                                                 date, source, stence]

            ##################################
            # ONE OR MORE PEOPLE REFERENCED
            ##################################
            elif len(prows)>0:
                # Since no said equivalent, but one or more people referenced,
                # check for quotes because if quoted then the most recent person referenced should be the speaker and the person/people
                # referenced in the sentence would either be inside a quote as in '"It was such a great year for Steve Jobs."')
                # or just in the sentence as in "'We caught up with him and he shared his thoughts about Steve Jobs: "What a great year he had."

                x1 = re.findall(r'(' + start + r')(.+)(' + end + r')', stence, re.IGNORECASE)
                x1 = ''.join(''.join(elems) for elems in x1)
                x1 = re.sub(r'[' + qs + r']', '', x1)



                # if so, assume must be from the most recent person referenced in the article
                if len(x1) > 0:
                    d = []
                    #check if quote is within "person"
                    for p in fulls:
                        t = re.findall(r'(' + start + r')(.+)(' + end + r')', p, re.IGNORECASE)
                        if len(t)>0:
                            d.append(t)
                    if len(d)==0:
                        quote = x1
                        quote = re.sub(r'^' + stext + r'.+', '', quote)
                        quote = re.sub(r'' + stext + r'\Z', '', quote)
                        last = "NA"
                        person = "NA"
                        firm = "NA"
                        role = "NA"
                        qtype = "single sentence quote"
                        qsaid = "no"
                        qpref = "one or more"
                        comment = "Review, whether quote or something else, not clear"
                        df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment, AN, date,
                                                 source, stence]
                # Because said equivalent but no full quote, check if it's a beginning of multi-sentence quote
                # If so, assume it would have to be from the most recent referenced person
                # As in "he said", referring to the most recent person
                elif len(x1) == 0:
                    # check if a quote begins in this sentence:
                    q1 = re.findall(r'(' + start + r')(.+)', stence, re.IGNORECASE)
                    q1 = ''.join(''.join(elems) for elems in q1)
                    q1 = re.sub(r'[' + qs + r']', '', q1)

                    # If a quote does begin in this sentence, concatinate the subsequent sentences until the quote ends
                    if len(q1) > 0:
                        for s in range(n, len(doc.sentences)):
                            j = doc.sentences[s]
                            t = j.text
                            q2 = re.findall(r'(.+)[' + end + r']', t, re.IGNORECASE)
                            q2 = ' '.join(q2)
                            q2 = re.sub(r'[' + qs + r']', '', q2)
                            if len(q2) == 0:
                                q1 = q1 + " " + t
                            if len(q2) > 0:
                                q1 = q1 + " " + q2
                                break
                        quote = q1
                        quote = re.sub(r'^' + stext + r'.+', '', quote)
                        quote = re.sub(r'' + stext + r'\Z', '', quote)
                        last = sppl[-1][0]
                        person = sppl[-1][1]
                        firm = sppl[-1][2]
                        role = sppl[-1][3]
                        qtype = "multi-sentence quote"
                        qsaid = "no"
                        qpref = "one or more"
                        comment = "high accuracy - prior person referenced"
                        df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment, AN,
                                                 date, source, stence]
                    # If there's no multi-sentence beginning here and no full quote, then it's just information with reference to 1+ people
                    elif len(q1) == 0:
                        for per in prows:
                            quote = stence
                            last = per[0]
                            person = per[1]
                            firm = per[2]
                            role = per[3]
                            qtype = "Information"
                            qsaid = "no"
                            qpref = "one or more"
                            comment = "Just information concerning one or more people"
                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment,
                                                     AN, date, source, stence]

    ##################################
    # NO PEOPLE IN CUMULATIVE LIST YET
    ##################################
    if len(people) == 0:
        # This should be rare because typically the first mentioning of a person includes the title and organization.
        # Because no person (i.e., with firm, title, and name) associated with it yet, see if at least a Person:
        # If exactly 1 person, check if quotes. No need to check for pronouns yet because no reference yet:

        if len(ppl) == 1:

            # Quote then person then said equivalent
            s1a = re.findall(r'(' + start + r')(.+)(' + end + r').+(' + per + r').+(' + stext + r')', stence,
                             re.IGNORECASE)
            s1b = re.findall(r'(' + start + r')(.+)(' + end + r').+(' + stext + r').+(' + per + r')', stence,
                             re.IGNORECASE)
            # Said equiv then last name then a quote (distance not matter here because sometimes full title in between)
            s1c = re.findall(r'(' + per + r').+(' + stext + r').+(' + start + r')(.+)(' + end + r')', stence,
                             re.IGNORECASE)
            s1d = re.findall(r'(' + stext + r').+(' + per + r').+(' + start + r')(.+)(' + end + r')', stence,
                             re.IGNORECASE)
            s1 = s1a + s1b + s1c + s1d
            s1 = ''.join(''.join(elems) for elems in s1)
            s1 = re.sub(r'[' + qs + r']', '', s1)
            s1 = re.sub(r'(' + stext + r')', '', s1)
            s1 = re.sub(r'' + per + r'', '', s1)
            if len(s1) > 0:
                person = per
                last = per.split()[-1]
                firm = "NA"
                role = "NA"
                qtype = "single sentence quote"
                qsaid = "yes"
                qpref = "yes"
                comment = "review - no full profile of person yet"
                quote = s1
                quote = re.sub(r'^' + stext + '.+', '', quote)
                quote = re.sub(r'' + stext + '\Z', '', quote)
                quote = re.sub(r'' + per + '', '', quote)
                df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment, AN, date,
                                         source, stence]
            elif len(s1) == 0:
                # check if a quote begins in this sentence:
                q1a = re.findall(r'(' + per + r').+(' + stext + r').+(' + start + r')(.+)', stence, re.IGNORECASE)
                q1b = re.findall(r'(' + stext + r').+(' + per + r').+(' + start + r')(.+)', stence, re.IGNORECASE)
                q1 = q1a + q1b
                q1 = ''.join(''.join(elems) for elems in q1)
                q1 = re.sub(r'[' + qs + r']', '', q1)
                q1 = re.sub(r'(' + stext + r')', '', q1)
                q1 = re.sub(r'' + per + r'', '', q1)

                # If a quote does begin in this sentence, concatinate the subsequent sentences until the quote ends
                if len(q1) > 0:
                    for s in range(n, len(doc.sentences)):
                        j = doc.sentences[s]
                        t = j.text
                        q2 = re.findall(r'(.+)[' + end + r']', t, re.IGNORECASE)
                        q2 = ' '.join(q2)
                        q2 = re.sub(r'[' + qs + r']', '', q2)
                        if len(q2) == 0:
                            q1 = q1 + " " + t
                        if len(q2) > 0:
                            q1 = q1 + " " + q2
                            break

                    last = "NA"
                    person = per
                    firm = "NA"
                    role = "NA"
                    qtype = "multi-sentence quote"
                    qsaid = "yes"
                    qpref = "yes"
                    quote = q1
                    quote = re.sub(r'^' + stext + r'.+', '', quote)
                    quote = re.sub(r'' + stext + r'\Z', '', quote)
                    comment = "review - no full profile of person yet"
                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment, AN, date,
                                             source, stence]
                #If no people in cumulative list or in this
                elif len(q1) == 0:
                    if len(orgs) > 1:
                        s = re.findall(r'' + stext + '', stence, re.IGNORECASE)
                        if len(s)>0:
                            last = "NA"
                            person = "NA"
                            firm = orgs[0]
                            role = "NA"
                            qtype = "paraphrase"
                            qsaid = "yes"
                            qpref = "one firm"
                            comment = "statement by the firm"
                            quote = stence
                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment, AN,
                                                     date, source, stence]
                        elif len(orgs) > 1:
                            s = re.findall(r'' + stext + '', stence, re.IGNORECASE)
                            if len(s) > 0:
                                last = "NA"
                                person = "NA"
                                firm = ''.join(''.join(elems) for elems in orgs)
                                role = "NA"
                                qtype = "paraphrase"
                                qsaid = "yes"
                                qpref = "multiple firms"
                                comment = "review - statement by one of multiple firms"
                                quote = stence
                                df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment, AN,
                                                         date, source, stence]

        elif len(ppl) > 1:
            # Very rare. If more than one person (without titles/orgs), include any organizations in the sentence
            # flag for review
            for per in ppl:
                quote = stence
                last = per.split()[-1]
                person = per
                firm = ''.join(''.join(elems) for elems in orgs)
                role = "NA"
                qtype = "Information"
                qsaid = "no"
                qpref = "more than one"
                comment = "review - no full profile of person yet"
                df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment, AN, date,
                                         source, stence]
        elif len(ppl) == 0:
            if len(orgs) == 1:
                s = re.findall(r'' + stext + '', stence, re.IGNORECASE)
                if len(s)>0:
                    last = "NA"
                    person = "NA"
                    firm = orgs[0]
                    role = "NA"
                    qtype = "paraphrase"
                    qsaid = "yes"
                    qpref = "one - firm"
                    comment = "statement by the firm"
                    quote = stence
                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment, AN, date,
                                             source, stence]
                elif len(s)==1:
                    last = "NA"
                    person = "NA"
                    firm = orgs[0]
                    role = "NA"
                    qtype = "information"
                    qsaid = "yes"
                    qpref = "one - firm"
                    comment = "info about a firm"
                    quote = stence
                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment, AN, date,
                                             source, stence]

            elif len(orgs) > 1:
                s = re.findall(r'' + stext + '', stence, re.IGNORECASE)
                if len(s)>0:
                    #If someone said something, see if it's a representative of a firm
                    for org in orgs:
                        m1 = re.findall(r'' + org + r'\s+(' + rep + r')', stence, re.IGNORECASE)
                        m2 = re.findall(r'(' + rep + r')\s+of\s+' + org + r'', stence, re.IGNORECASE)
                        m = m1 + m2
                        m = ''.join(''.join(elems) for elems in m)
                        m = re.sub(r'(' + org + r')', '', m)
                        if len(m)>0:
                            last = "NA"
                            person = m
                            firm = org
                            role = "NA"
                            qtype = "paraphrase"
                            qsaid = "yes"
                            qpref = "one or more firms"
                            comment = "one or more firm representatives said something"
                            quote = stence
                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment,
                                                     AN, date,
                                                     source, stence]

    else:
        pass

    # for ent in sent.ents:
    #    print (n, ent.text, ent.type)

    # print (n, "PERSON:",person,"ROLE:",role,"FIRM:",firm, "QUOTE:",quote, "TYPE:",qtype,"SAID:",qsaid,"REF:",qpref,"COMMENTS:",comment)
    # print(n, people)

    print (n,stence)
    n += 1

df

#df.to_csv(r'C:\Users\danwilde\Dropbox (Penn)\Dissertation\Factiva\data1.csv')

# print(people)

# TO DO:
# Interviews
# Analyst calls
# Should add any information options if no 'people' but firm or person?



