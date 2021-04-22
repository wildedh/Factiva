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

dfa = pd.read_excel(r'C:\Users\danwilde\Dropbox (Penn)\Dissertation\Factiva\final7.xlsx')

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

# combine the lead paragraph and the body of the article

columns = ['sent', 'person', 'last', 'role', 'firm', 'quote', 'qtype', 'qsaid', 'qpref', 'comment', 'AN', 'date',
           'source', 'regions', 'subjects', 'misc', 'industries', 'focfirm', 'by', 'comp', 'sentiment', 'sentence']

df = pd.DataFrame(columns=columns)

stext = r'said|n.t say|say|told|n.t tell\S+|tell\S+|assert\S+|according\s+to|n.t comment\S+|comment\S+|quote\S+' \
        r'|describ\S+|n.t communicat\S+|communicat\S+|articulat\S+|n.t divulg\S+|divulg\S+|noted|noting|espressed' \
        r'|recounted|suggested|\bexplained\b|\badded\b|acknowledged|\bstating\b|\bstated\b|\bprotested\b|\bfumed\b|recounts|\basks\b|\basked\b'

l1 = r'executive\s+vice\s+president\s+of\s\S+|executive\s+vice-president\s+of\s\S+|' \
     r'senior\s+vice\s+president\b\s+of|\S+\s+executive\b\s+of\s+the\s+\S+|vice\s+president\s+and\s+general\s+\S+'

l2 = r'executive\s+vice\s+president|executive\s+vice-president|sr\.\s+vice\s+president\b\s+of\s+\S|' \
     r'head\s+of\s\S+|chair\S+\s+and\s+\S+\s+head\b|senior\s+vice-president\b|' \
     r'senior\s+director\s+of\s+\S+|chief.+officer|president\s+and\s+C.0|chair\S+\s+and\s+\S+'

l3 = r'vice\s+chair\S+|vice\s+president|vice-president|senior\s+manager|sr.*manager|senior\s+director|' \
     r'chief\s+executive|senior\s+manager|sr.*manager|vp\s+of\s+\S+|global\+director|senior\s+director|' \
     r'sr.*director|managing\s+director|vice\s+minister|prime\s+minister|founding\s+partner'

l4 = r'president|chair\S+|director|manager|vp|' \
     r'banker\b|specialist\b|accountant\b|journalist\b|reporter\b|analyst\b|consultant\b|negotiator\b|' \
     r'governor\b|congress\S+|house\s+speaker|senator|senior\s+fellow|fellow|undersecretary|' \
     r'spokesman\b|spokesperson\b|spokeswoman\b|representative\b|official\b|executive\b|aide|pilot|professor|' \
     r'attorney|minister|secretary|guru|partner|general\s+counsel'

l5 = r'\bC.O\b|board|administrator'

pron = r'\bhe\b|\bshe\b'
plurals = r'analysts|bankers|consultants|executives|management|participants|officials'
rep = r'spokesman\b|spokesperson\b|spokeswoman\b|representative\b|official\b|executive\b'
start = r'\"|\'\'|\`\`'
end = r'\"|\'\''
qs = r'\`\`\"|\'\''
suf = r'\bJr\b|\bSr\b|\sI\s|\sII\s|\sIII\s|\bPhD\b|Ph\.D|\bMBA\b|\bCPA\b'
news = r'Reuters|'

z = 0

arts = [23]

for art in arts:
#for art in dfa.index:

    t1 = time.time()
    text = str(dfa['LP'][art]) + str(dfa['TD'][art])
    #text = str(dfa['LP'][art])
    AN = dfa["AN"][art]
    date = dfa["PD"][art]
    source = dfa["SN"][art]
    regions = dfa["RE"][art]
    subjects = dfa["NS"][art]
    misc = dfa["IPD"][art]
    industries = dfa["IN"][art]
    focfirm = dfa["Firm"][art]
    by = dfa["BY"][art]
    comp = dfa["CO"][art]
    ##############################################################
    # Prepocessing the text to let it be analyzed by the Stanza model.
    ##############################################################

    # First, need to remove the dateline: the brief piece of text included in news articles describing
    # where and when the story was written or filed. Keeping it can erroneiously code one extra
    # relevant org such as the source (e.g., Reuters) that would mess up the attribution of quotes
    # in the code below.

    # A few ways to identify a dateline:
    d1a = re.findall(r'\(.+\)\s*--', text, re.IGNORECASE)
    d1b = re.findall(r'\(.+\)\s*-', text, re.IGNORECASE)
    d2a = re.findall(r'^.+/\b.+\b/\s*--', text, re.IGNORECASE)
    d2b = re.findall(r'^.+/\b.+\b/\s*-', text, re.IGNORECASE)
    d3a = re.findall(r'^.+\d\d*\s*--', text, re.IGNORECASE)
    d3b = re.findall(r'^.+\d\d*\s-', text, re.IGNORECASE)

    # Iteratively go through these instances. Only do one because could potentially
    # over cut if run on more than one.

    if len(d1a) > 0:
        text = re.sub(r'^.+\(.+\)\s*--', '', text)

    elif len(d1b) > 0:
        text = re.sub(r'^.+\(.+\)\s*-', '', text)

    elif len(d2a) > 0:
        text = re.sub(r'^.+/\b.+\b/\s*--', '', text)

    elif len(d2b) > 0:
        text = re.sub(r'^.+/\b.+\b/\s*-', '', text)

    elif len(d3a) > 0:
        text = re.sub(r'^.+\d\d*\s--', '', text)

    elif len(d3b) > 0:
        text = re.sub(r'^.+\d\d*\s-', '', text)


    # Next, remove any parentheses. These mess up the NLP process as they mean something literally
    # in the PyTorch code.
    text = re.sub(r'\(|\)', '', text)
    
    # Next remove any quotes around a single word (e.g., "great"). These also irratate the nlp program
    zy = []
    t = re.findall(r'(' + start + r')(\w+\w)(' + end +r')',text)
    for w in t:
        u = ''.join(''.join(elems) for elems in w)
        y = re.sub(r'[' + qs + r']','', u)
        z = [u,y]
        zy.append(z)
    for w in zy:
        text = re.sub(r'(' + w[0] + r')',w[1],text)
    
    # Next replace multi-dashed words with no dashes.
    v = re.findall(r'\w+-\w+-\w+',text)
    for w in v:
        u = ''.join(''.join(elems) for elems in w)
        y = re.sub(r'-', '', u)
        z = [u, y]
        zy.append(z)
    for w in zy:
        text = re.sub(r'(' + w[0] + r')', w[1], text)


    doc = nlp(text)

    people = []
    sppl = []

    fulls = []
    lsts = []

    corgs = []

    # Set the max number of words away from the focal leader and the focal organization
    mxld = 6
    mxorg = 6
    mxlds = 2
    mxpron = 2
    mxorg1 = 2

    n = 1
    dx = 0



    #try:




    # Look through document and see if has instances of at least four all-cap words
    # broken up by non-word characters (e.g., spaces, commas) followed by a colon.
    f = re.findall(r'(((\b[A-Z]+\b\W*){4,})\:)', text)
    # If there is at least 1 of these very unique instances, this is got to be a transcript of some kind.
    if len(f) > 0:
        # Go through each sentence to find if it's introducing a person and their role and firm
        for sent in doc.sentences:

            stence = sent.text
            # Make sure there is something in the sentence. If not, skip.
            op = re.findall(r'^OPERATOR', stence)
            if len(op) == 0:

                # Set a trigger variable to 0 once start a new sentence.
                dr = 0
                # Calculate the sentiment for each sentence
                ment = sent.sentiment

                #for ent in sent.ents:
                #    print(n, ent.text, ent.type)

                last = ""
                role = ""
                person = ""
                firm = ""
                # Check if name title in the sentence:
                g = re.findall(r'(((\b[A-Z]+\b\W*){4,})\:)', stence)
                if len(g) > 0:
                    a = g[0][0]
                    a = re.sub(r':', '', a)
                    # If company name is Inc and has a comma, remove the comma
                    a = re.sub('\,\s+(INC|Inc)', ' Inc', a, re.IGNORECASE)
                    p = a.split(',')

                    person = p[0]
                    # Figure out the last name depending on if there's a suffix
                    sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                    if len(sfx) > 0:
                        last = person.split()[-2]
                    elif len(sfx) == 0:
                        last = person.split()[-1]
                    firm = p[-1]
                    # Lastly, if there are four pieces the role is the 2nd and 3rd pieces so combine
                    if len(p) == 4:
                        r = p[1] + p[2]
                        role = ''.join(''.join(elems) for elems in r)
                    elif len(p) == 3:
                        role = p[1]
                    row = [last, person, firm, role, n]

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
                        h = re.findall(r'' + last + r':', stence, re.IGNORECASE)
                        m = m + h
                    m = ''.join(''.join(elems) for elems in m)
                    if len(m) > 0:
                        for per in people:
                            last = per[0]
                            l = re.findall(r'' + last + r':', stence, re.IGNORECASE)
                            if len(l) > 0:
                                # Strip the beginning with the quote.
                                quote = re.sub(r'^.*' + last + r':', '', stence)
                                last = per[0]
                                person = per[1]
                                firm = per[2]
                                role = per[3]
                                qtype = "Transcript"
                                qsaid = "yes"
                                qpref = "yes"
                                comment = "Transcript"
                                df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref,
                                                         comment,
                                                         AN, date, source, regions, subjects, misc, industries,
                                                         focfirm, by, comp, ment, stence]
                                dx += 1
                                # Add this person to the running list of sequential people speaking
                                sppl.append(per)
                                break


                    # If there isn't an introduction of a person or a reintroduction of a person
                    elif len(m) == 0:
                        # Are there any people in the people list yet?
                        # If so, then the sentence belongs to the last person assigned.
                        if len(people) > 0:
                            quote = re.sub(r'^.*' + last + r':', '', stence)
                            last = sppl[-1][0]
                            person = sppl[-1][1]
                            firm = sppl[-1][2]
                            role = sppl[-1][3]
                            qtype = "Transcript"
                            qsaid = "yes"
                            qpref = "yes"
                            comment = "Transcript"
                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref,
                                                     comment,
                                                     AN, date,
                                                     source, regions, subjects, misc, industries, focfirm, by, comp,
                                                     ment, stence]
                            dx += 1
                n += 1


    #####################
    # Non-transcripts
    #####################
    elif len(f) == 0:

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

            stence = sent.text

            # Calculate the sentiment for each sentence
            ment = sent.sentiment

            #for ent in sent.ents:
            #    print(n, ent.text, ent.type)

            # Must stip the the sentence of any quotes in order to arrive at appropriate orgs, people, and products
            # First, remove any information within quotes:
            stz1 = re.sub(r'(' + start + r').+(' + end + r')', ' ', stence)
            # Second, remove information in beginning of multi-sentence quotes
            # the start quote would have a space before the quote
            stz2 = re.sub(r'([:|,])\s+(' + start + r')\S.+$', ' ', stz1)
            # Third, remove cases in which the quote begins in the sentance and doesn't end.
            stz3 = re.sub(r'^(' + start + r')\S.+$', ' ', stz2)
            # Last, remove information at the end of multi-sentence quotes
            # the end quote would have a space after the quote
            stz4 = re.sub(r'^.+(' + end + r')\s+', ' ', stz3)

            # Check if you can add to dataset of people/firm/role
            # First, compile a list of people and organizations and products iteratively within the sentence.
            # You use products list in the orgs list so need to run them sep.
            for ent in sent.ents:
                if ent.type == "PERSON":
                    z = re.findall(r'' + ent.text + r'', stz4)
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
                    pduct = re.sub(r'\(|\)', '', pduct)
                    z = re.findall(r'' + pduct + r'', stz4)
                    if len(z) > 0:
                        prod.append(pduct)
            # Finally, the orgs
            for ent in sent.ents:
                if ent.type == "ORG":
                    tion = ent.text
                    # remove odd punctuation such as ")" that can mess up the regex down the line
                    tion = re.sub(r'\(|\)|\-', '', tion)
                    z = re.findall(r'' + tion + r'', stz4)
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
            # First check if just one person
            # Second check if just more than one person
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
                            pos1a = re.findall(r'(' + per + r').+(' + l1 + r')', stz4, re.IGNORECASE)
                            pos1b = re.findall(r'(' + per + r').+(' + l2 + r')', stz4, re.IGNORECASE)
                            pos1c = re.findall(r'(' + per + r').+(' + l3 + r')', stz4, re.IGNORECASE)
                            pos1d = re.findall(r'(' + per + r').+(' + l4 + r')', stz4, re.IGNORECASE)
                            pos1e = re.findall(r'(' + per + r').+(' + l5 + r')', stz4)
                            pos2a = re.findall(r'(' + l1 + r').+(' + per + r')', stz4, re.IGNORECASE)
                            pos2b = re.findall(r'(' + l2 + r').+(' + per + r')', stz4, re.IGNORECASE)
                            pos2c = re.findall(r'(' + l3 + r').+(' + per + r')', stz4, re.IGNORECASE)
                            pos2d = re.findall(r'(' + l4 + r').+(' + per + r')', stz4, re.IGNORECASE)
                            pos2e = re.findall(r'(' + l5 + r').+(' + per + r')', stz4)
                            pos = pos1a + pos1b + pos1c + pos1d + pos1e + pos2a + pos2b + pos2c + pos2d + pos2e

                        # If either 0 or >1 orgs, do this:
                        else:
                            pos1a = re.findall(
                                r'\b(?:(' + per + r')\W+(?:\w+\W+){0,' + str(mxld) + r'}?(' + l1 + r'))',
                                stz4,
                                re.IGNORECASE)
                            pos1b = re.findall(
                                r'\b(?:(' + per + r')\W+(?:\w+\W+){0,' + str(mxld) + r'}?(' + l2 + r'))',
                                stz4,
                                re.IGNORECASE)
                            pos1c = re.findall(
                                r'\b(?:(' + per + r')\W+(?:\w+\W+){0,' + str(mxld) + r'}?(' + l3 + r'))',
                                stz4,
                                re.IGNORECASE)
                            pos1d = re.findall(
                                r'\b(?:(' + per + r')\W+(?:\w+\W+){0,' + str(mxld) + r'}?(' + l4 + r'))',
                                stz4,
                                re.IGNORECASE)
                            pos1e = re.findall(
                                r'\b(?:(' + per + r')\W+(?:\w+\W+){0,' + str(mxld) + r'}?(' + l5 + r'))',
                                stz4)
                            pos2a = re.findall(
                                r'\b(?:(' + l1 + r')\W+(?:\w+\W+){0,' + str(mxld) + r'}?(' + per + r'))',
                                stz4,
                                re.IGNORECASE)
                            pos2b = re.findall(
                                r'\b(?:(' + l2 + r')\W+(?:\w+\W+){0,' + str(mxld) + r'}?(' + per + r'))',
                                stz4,
                                re.IGNORECASE)
                            pos2c = re.findall(
                                r'\b(?:(' + l3 + r')\W+(?:\w+\W+){0,' + str(mxld) + r'}?(' + per + r'))',
                                stz4,
                                re.IGNORECASE)
                            pos2d = re.findall(
                                r'\b(?:(' + l4 + r')\W+(?:\w+\W+){0,' + str(mxld) + r'}?(' + per + r'))',
                                stz4,
                                re.IGNORECASE)
                            pos2e = re.findall(
                                r'\b(?:(' + l5 + r')\W+(?:\w+\W+){0,' + str(mxld) + r'}?(' + per + r'))',
                                stz4)
                            pos = pos1a + pos1b + pos1c + pos1d + pos1e + pos2a + pos2b + pos2c + pos2d + pos2e


                    # Then by definition, must be >1 people
                    else:
                        pos1a = re.findall(
                            r'\b(?:(' + per + r')\W+(?:\w+\W+){0,' + str(mxlds) + r'}?(' + l1 + r'))',
                            stz4,
                            re.IGNORECASE)
                        pos1b = re.findall(
                            r'\b(?:(' + per + r')\W+(?:\w+\W+){0,' + str(mxlds) + r'}?(' + l2 + r'))',
                            stz4,
                            re.IGNORECASE)
                        pos1c = re.findall(
                            r'\b(?:(' + per + r')\W+(?:\w+\W+){0,' + str(mxlds) + r'}?(' + l3 + r'))',
                            stz4,
                            re.IGNORECASE)
                        pos1d = re.findall(
                            r'\b(?:(' + per + r')\W+(?:\w+\W+){0,' + str(mxlds) + r'}?(' + l4 + r'))',
                            stz4,
                            re.IGNORECASE)
                        pos1e = re.findall(
                            r'\b(?:(' + per + r')\W+(?:\w+\W+){0,' + str(mxlds) + r'}?(' + l5 + r'))',
                            stz4)
                        pos2a = re.findall(
                            r'\b(?:(' + l1 + r')\W+(?:\w+\W+){0,' + str(mxlds) + r'}?(' + per + r'))',
                            stz4,
                            re.IGNORECASE)
                        pos2b = re.findall(
                            r'\b(?:(' + l2 + r')\W+(?:\w+\W+){0,' + str(mxlds) + r'}?(' + per + r'))',
                            stz4,
                            re.IGNORECASE)
                        pos2c = re.findall(
                            r'\b(?:(' + l3 + r')\W+(?:\w+\W+){0,' + str(mxlds) + r'}?(' + per + r'))',
                            stz4,
                            re.IGNORECASE)
                        pos2d = re.findall(
                            r'\b(?:(' + l4 + r')\W+(?:\w+\W+){0,' + str(mxlds) + r'}?(' + per + r'))',
                            stz4,
                            re.IGNORECASE)
                        pos2e = re.findall(
                            r'\b(?:(' + l5 + r')\W+(?:\w+\W+){0,' + str(mxlds) + r'}?(' + per + r'))',
                            stz4)
                        pos = pos1a + pos1b + pos1c + pos1d + pos1e + pos2a + pos2b + pos2c + pos2d + pos2e

                    # If the title is mentioned, move to next step
                    if len(pos) > 0:

                        # If there is just one organization mentioned, assign
                        if len(orgs) == 1:
                            org = orgs[0]
                            org = org.replace(r"'s", "")
                            org1 = re.findall(r'(' + per + r').+(' + org + r')', stz4, re.IGNORECASE)
                            org2 = re.findall(r'(' + org + r').+(' + per + r')', stz4, re.IGNORECASE)
                            f = org1 + org2
                            f = ''.join(''.join(elems) for elems in f)

                            # If yes, add as potential person
                            if len(f) > 0:
                                person = per
                                sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                                if len(sfx) > 0:
                                    last = person.split()[-2]
                                elif len(sfx) == 0:
                                    last = person.split()[-1]

                                role = ""
                                for i in pos:
                                    j = ''.join(''.join(elems) for elems in i)
                                    role = re.sub(r'' + per + r'', '', j)
                                    role = re.sub(r',', '', role)
                                    if len(role) > 0:
                                        break
                                f = re.sub(r'(' + per + r')', '', f)
                                firm = f
                        # Multiple orgs in sentence
                        elif len(orgs) > 1:

                            # Create list and variable that will be used at the end of this iteration of orgs
                            # "gm" checks if there are any instances in which an org is a substring of another
                            #  org name. If yes, len(g)>0. If not, len(g) = 0 and you're good.
                            c = 0
                            b = 0
                            h = 0
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








                            for o in orgs:
                                # We want to make sure that the right org name gets assigned. The problem is
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

                                if c > 0 and len(b) > 0:
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

                                    if h>0 and l==0:
                                        continue


                                    # Check if there is possession with one "'s" in the org
                                    o1 = re.findall(r'' + o + r'\'s', stz4, re.IGNORECASE)
                                    o2 = re.findall(r'\'s', o, re.IGNORECASE)

                                    og = o1 + o2


                                    if len(og) > 0:
                                        o1a = re.findall(
                                            r'\b(?:(' + o + r')\W+(?:\w+\W+){0,' + str(
                                                mxorg1) + r'}?(' + l1 + r'))',
                                            stz4, re.IGNORECASE)
                                        o1b = re.findall(
                                            r'\b(?:(' + o + r')\W+(?:\w+\W+){0,' + str(
                                                mxorg1) + r'}?(' + l2 + r'))',
                                            stz4, re.IGNORECASE)
                                        o1c = re.findall(
                                            r'\b(?:(' + o + r')\W+(?:\w+\W+){0,' + str(
                                                mxorg1) + r'}?(' + l3 + r'))',
                                            stz4, re.IGNORECASE)
                                        o1d = re.findall(
                                            r'\b(?:(' + o + r')\W+(?:\w+\W+){0,' + str(
                                                mxorg1) + r'}?(' + l4 + r'))',
                                            stz4, re.IGNORECASE)
                                        o1e = re.findall(
                                            r'\b(?:(' + o + r')\W+(?:\w+\W+){0,' + str(
                                                mxorg1) + r'}?(' + l5 + r'))',
                                            stz4)
                                        o2a = re.findall(
                                            r'\b(?:(' + l1 + r')\W+(?:\w+\W+){0,' + str(
                                                mxorg1) + r'}?(' + o + r'))',
                                            stz4, re.IGNORECASE)
                                        o2b = re.findall(
                                            r'\b(?:(' + l2 + r')\W+(?:\w+\W+){0,' + str(
                                                mxorg1) + r'}?(' + o + r'))',
                                            stz4, re.IGNORECASE)
                                        o2c = re.findall(
                                            r'\b(?:(' + l3 + r')\W+(?:\w+\W+){0,' + str(
                                                mxorg1) + r'}?(' + o + r'))',
                                            stz4, re.IGNORECASE)
                                        o2d = re.findall(
                                            r'\b(?:(' + l4 + r')\W+(?:\w+\W+){0,' + str(
                                                mxorg1) + r'}?(' + o + r'))',
                                            stz4, re.IGNORECASE)
                                        o2e = re.findall(
                                            r'\b(?:(' + l5 + r')\W+(?:\w+\W+){0,' + str(
                                                mxorg1) + r'}?(' + o + r'))',
                                            stz4)
                                        os = o1a + o1b + o1c + o1d + o1e + o2a + o2b + o2c + o2d + o2e
                                        om = ''.join(''.join(elems) for elems in os)

                                        if len(os) > 0:
                                            person = per
                                            sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                                            if len(sfx) > 0:
                                                last = person.split()[-2]
                                            elif len(sfx) == 0:
                                                last = person.split()[-1]
                                            role = ""
                                            for i in os:
                                                j = ''.join(''.join(elems) for elems in i)
                                                role = re.sub(r'' + o + r'', '', j)
                                                role = re.sub(r',', '', role)
                                                if len(role) > 0:
                                                    break
                                            o = o.replace(r"'s", "")
                                            firm = o

                                            c += 1
                                            h += 1
                                            if len(gm) == 0:
                                                break

                                    # If not, check if there is possession with "whose"
                                    # Org then whose then title.
                                    elif len(og) == 0:

                                        o3a = re.findall(
                                            r'\b(?:(' + o + r')\W+(?:\w+\W+){0,2}?(whose\s+(' + l1 + ')))',
                                            stz4,
                                            re.IGNORECASE)
                                        o3b = re.findall(
                                            r'\b(?:(' + o + r')\W+(?:\w+\W+){0,2}?(whose\s+(' + l2 + ')))',
                                            stz4,
                                            re.IGNORECASE)
                                        o3c = re.findall(
                                            r'\b(?:(' + o + r')\W+(?:\w+\W+){0,2}?(whose\s+(' + l3 + ')))',
                                            stz4,
                                            re.IGNORECASE)
                                        o3d = re.findall(
                                            r'\b(?:(' + o + r')\W+(?:\w+\W+){0,2}?(whose\s+(' + l4 + ')))',
                                            stz4,
                                            re.IGNORECASE)
                                        o3e = re.findall(
                                            r'\b(?:(' + o + r')\W+(?:\w+\W+){0,2}?(whose\s+(' + l5 + ')))',
                                            stz4)

                                        m = o3a + o3b + o3c + o3d + o3e

                                        role = ""

                                        # If whose then title, assign
                                        if len(m) > 0:
                                            person = per
                                            sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                                            if len(sfx) > 0:
                                                last = person.split()[-2]
                                            elif len(sfx) == 0:
                                                last = person.split()[-1]

                                            role = ""
                                            for i in m:
                                                # If match, there are three items, the third item (index 2) is the role
                                                j = i[2]
                                                role = re.sub(r'' + o + r'', '', j)
                                                role = re.sub(r',', '', role)
                                                if len(role) > 0:
                                                    break
                                            firm = o
                                            c += 1
                                            h += 1
                                            if len(gm) == 0:
                                                break

                                        # If not, check if [title] "of" near [org]
                                        elif len(m) == 0:


                                            z1 = re.findall(
                                                r'\b(?:(' + l1 + r')(\s+of)\W+(?:\w+\W+){0,2}?(' + o + '))', stz4,
                                                re.IGNORECASE)
                                            z2 = re.findall(
                                                r'\b(?:(' + l2 + r')(\s+of)\W+(?:\w+\W+){0,2}?(' + o + '))', stz4,
                                                re.IGNORECASE)
                                            z3 = re.findall(
                                                r'\b(?:(' + l3 + r')(\s+of)\W+(?:\w+\W+){0,2}?(' + o + '))', stz4,
                                                re.IGNORECASE)
                                            z4 = re.findall(
                                                r'\b(?:(' + l4 + r')(\s+of)\W+(?:\w+\W+){0,2}?(' + o + '))', stz4,
                                                re.IGNORECASE)
                                            z5 = re.findall(
                                                r'\b(?:(' + l5 + r')(\s+of)\W+(?:\w+\W+){0,2}?(' + o + '))', stz4)
                                            m = z1 + z2 + z3 + z4 + z5

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
                                                h += 1
                                                if len(gm) == 0:
                                                    break


                                            # If not any of these three strong matches ('s, of, whose),
                                            # check on a weaker match: if name and person are close to title
                                            elif len(j) == 0:

                                                o5a = re.findall(
                                                    r'\b(?:(' + o + r')\W+(?:\w+\W+){0,' + str(
                                                        mxorg1) + r'}?(' + l1 + r'))',
                                                    stz4, re.IGNORECASE)
                                                o5b = re.findall(
                                                    r'\b(?:(' + o + r')\W+(?:\w+\W+){0,' + str(
                                                        mxorg1) + r'}?(' + l2 + r'))',
                                                    stz4, re.IGNORECASE)
                                                o5c = re.findall(
                                                    r'\b(?:(' + o + r')\W+(?:\w+\W+){0,' + str(
                                                        mxorg1) + r'}?(' + l3 + r'))',
                                                    stz4, re.IGNORECASE)
                                                o5d = re.findall(
                                                    r'\b(?:(' + o + r')\W+(?:\w+\W+){0,' + str(
                                                        mxorg1) + r'}?(' + l4 + r'))',
                                                    stz4, re.IGNORECASE)
                                                o5e = re.findall(
                                                    r'\b(?:(' + o + r')\W+(?:\w+\W+){0,' + str(
                                                        mxorg1) + r'}?(' + l5 + r'))',
                                                    stz4)
                                                o6a = re.findall(
                                                    r'\b(?:(' + l1 + r')\W+(?:\w+\W+){0,' + str(
                                                        mxorg1) + r'}?(' + o + r'))',
                                                    stz4, re.IGNORECASE)
                                                o6b = re.findall(
                                                    r'\b(?:(' + l2 + r')\W+(?:\w+\W+){0,' + str(
                                                        mxorg1) + r'}?(' + o + r'))',
                                                    stz4, re.IGNORECASE)
                                                o6c = re.findall(
                                                    r'\b(?:(' + l3 + r')\W+(?:\w+\W+){0,' + str(
                                                        mxorg1) + r'}?(' + o + r'))',
                                                    stz4, re.IGNORECASE)
                                                o6d = re.findall(
                                                    r'\b(?:(' + l4 + r')\W+(?:\w+\W+){0,' + str(
                                                        mxorg1) + r'}?(' + o + r'))',
                                                    stz4, re.IGNORECASE)
                                                o6e = re.findall(
                                                    r'\b(?:(' + l5 + r')\W+(?:\w+\W+){0,' + str(
                                                        mxorg1) + r'}?(' + o + r'))',
                                                    stz4)
                                                m = o5a + o5b + o5c + o5d + o5e + o6a + o6b + o6c + o6d + o6e

                                                on = ''.join(''.join(elems) for elems in m)

                                                # If so, check if person is also near the title
                                                if len(on) > 0:

                                                    o7a = re.findall(
                                                        r'\b(?:(' + per + r')\W+(?:\w+\W+){0,' + str(
                                                            mxorg1) + r'}?(' + l1 + r'))',
                                                        stz4, re.IGNORECASE)
                                                    o7b = re.findall(
                                                        r'\b(?:(' + per + r')\W+(?:\w+\W+){0,' + str(
                                                            mxorg1) + r'}?(' + l2 + r'))',
                                                        stz4, re.IGNORECASE)
                                                    o7c = re.findall(
                                                        r'\b(?:(' + per + r')\W+(?:\w+\W+){0,' + str(
                                                            mxorg1) + r'}?(' + l3 + r'))',
                                                        stz4, re.IGNORECASE)
                                                    o7d = re.findall(
                                                        r'\b(?:(' + per + r')\W+(?:\w+\W+){0,' + str(
                                                            mxorg1) + r'}?(' + l4 + r'))',
                                                        stz4, re.IGNORECASE)
                                                    o7e = re.findall(
                                                        r'\b(?:(' + per + r')\W+(?:\w+\W+){0,' + str(
                                                            mxorg1) + r'}?(' + l5 + r'))',
                                                        stz4)
                                                    o8a = re.findall(
                                                        r'\b(?:(' + l1 + r')\W+(?:\w+\W+){0,' + str(
                                                            mxorg1) + r'}?(' + per + r'))',
                                                        stz4, re.IGNORECASE)
                                                    o8b = re.findall(
                                                        r'\b(?:(' + l2 + r')\W+(?:\w+\W+){0,' + str(
                                                            mxorg1) + r'}?(' + per + r'))',
                                                        stz4, re.IGNORECASE)
                                                    o8c = re.findall(
                                                        r'\b(?:(' + l3 + r')\W+(?:\w+\W+){0,' + str(
                                                            mxorg1) + r'}?(' + per + r'))',
                                                        stz4, re.IGNORECASE)
                                                    o8d = re.findall(
                                                        r'\b(?:(' + l4 + r')\W+(?:\w+\W+){0,' + str(
                                                            mxorg1) + r'}?(' + per + r'))',
                                                        stz4, re.IGNORECASE)
                                                    o8e = re.findall(
                                                        r'\b(?:(' + l5 + r')\W+(?:\w+\W+){0,' + str(
                                                            mxorg1) + r'}?(' + per + r'))',
                                                        stz4)

                                                    m = o7a + o7b + o7c + o7d + o7e + o8a + o8b + o8c + o8d + o8e

                                                    ol = ''.join(''.join(elems) for elems in m)
                                                    # If name and person are close to title, assign
                                                    if len(ol) > 0:
                                                        person = per
                                                        sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                                                        if len(sfx) > 0:
                                                            last = person.split()[-2]
                                                        elif len(sfx) == 0:
                                                            last = person.split()[-1]

                                                        role = ""
                                                        for i in m:
                                                            j = ''.join(''.join(elems) for elems in i)
                                                            role = re.sub(r'' + per + r'', '', j)
                                                            role = re.sub(r',', '', role)
                                                            if len(role) > 0:
                                                                break
                                                        firm = o
                                                        c += 1
                                                        if len(gm) == 0:
                                                            break
                                                # If not, check if name followed by just one org and a title.
                                                elif len(on) == 0:

                                                    c1 = re.findall(r'' + per + r'.+(' + l1 + r').+(' + o + r')', stz4,
                                                                    re.IGNORECASE)
                                                    c2 = re.findall(r'' + per + r'.+(' + l2 + r').+(' + o + r')', stz4,
                                                                    re.IGNORECASE)
                                                    c3 = re.findall(r'' + per + r'.+(' + l3 + r').+(' + o + r')', stz4,
                                                                    re.IGNORECASE)
                                                    c4 = re.findall(r'' + per + r'.+(' + l4 + r').+(' + o + r')', stz4,
                                                                    re.IGNORECASE)
                                                    c5 = re.findall(r'' + per + r'.+(' + l5 + r').+(' + o + r')', stz4)
                                                    c6 = re.findall(r'' + per + r'.+(' + o + r').+(' + l1 + r')', stz4,
                                                                    re.IGNORECASE)
                                                    c7 = re.findall(r'' + per + r'.+(' + o + r').+(' + l2 + r')', stz4,
                                                                    re.IGNORECASE)
                                                    c8 = re.findall(r'' + per + r'.+(' + o + r').+(' + l3 + r')', stz4,
                                                                    re.IGNORECASE)
                                                    c9 = re.findall(r'' + per + r'.+(' + o + r').+(' + l4 + r')', stz4,
                                                                    re.IGNORECASE)
                                                    c10 = re.findall(r'' + per + r'.+(' + o + r').+(' + l5 + r')', stz4)
                                                    m = c1 + c2 + c3 + c4 + c5 + c6 + c7 + c8 + c9 + c10

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

                                role = ""
                                for i in pos:
                                    j = ''.join(''.join(elems) for elems in i)
                                    role = re.sub(r'' + per + r'', '', j)
                                    role = re.sub(r',', '', role)
                                    if len(role) > 0:
                                        break
                                firm = df.loc[len(df.index) - 1].at['firm']
                    # If there's no leadership position, check if there is a person
                    # If there is and it says "of" something, add it to the people list
                    elif len(pos) == 0:
                        if len(ppl) == 1:
                            if len(orgs) > 0:
                                for org in orgs:
                                    q = re.findall(r'.+' + ppl[0] + r'\s+(of|at)\s+' + org + r'', stz4,
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
                    row = [last, person, firm, role, n]

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
                    o5a = re.findall(
                        r'\b(?:(' + o + r')\W+(?:\w+\W+){0,' + str(mxorg1) + r'}?(' + l1 + r'))',
                        stz4, re.IGNORECASE)
                    o5b = re.findall(
                        r'\b(?:(' + o + r')\W+(?:\w+\W+){0,' + str(mxorg1) + r'}?(' + l2 + r'))',
                        stz4, re.IGNORECASE)
                    o5c = re.findall(
                        r'\b(?:(' + o + r')\W+(?:\w+\W+){0,' + str(mxorg1) + r'}?(' + l3 + r'))',
                        stz4, re.IGNORECASE)
                    o5d = re.findall(
                        r'\b(?:(' + o + r')\W+(?:\w+\W+){0,' + str(mxorg1) + r'}?(' + l4 + r'))',
                        stz4, re.IGNORECASE)
                    o5e = re.findall(
                        r'\b(?:(' + o + r')\W+(?:\w+\W+){0,' + str(mxorg1) + r'}?(' + l5 + r'))',
                        stz4)
                    o6a = re.findall(
                        r'\b(?:(' + l1 + r')\W+(?:\w+\W+){0,' + str(mxorg1) + r'}?(' + o + r'))',
                        stz4, re.IGNORECASE)
                    o6b = re.findall(
                        r'\b(?:(' + l2 + r')\W+(?:\w+\W+){0,' + str(mxorg1) + r'}?(' + o + r'))',
                        stz4, re.IGNORECASE)
                    o6c = re.findall(
                        r'\b(?:(' + l3 + r')\W+(?:\w+\W+){0,' + str(mxorg1) + r'}?(' + o + r'))',
                        stz4, re.IGNORECASE)
                    o6d = re.findall(
                        r'\b(?:(' + l4 + r')\W+(?:\w+\W+){0,' + str(mxorg1) + r'}?(' + o + r'))',
                        stz4, re.IGNORECASE)
                    o6e = re.findall(
                        r'\b(?:(' + l5 + r')\W+(?:\w+\W+){0,' + str(mxorg1) + r'}?(' + o + r'))',
                        stz4)
                    m = o5a + o5b + o5c + o5d + o5e + o6a + o6b + o6c + o6d + o6e

                    on = ''.join(''.join(elems) for elems in m)

                    # If there is a role and org close together, then check if a single person in the next sentence:
                    if len(on) > 0:
                        pz = []
                        j = doc.sentences[n]
                        b = j.text
                        #### Filter out any quotes, then use that as the reference for stz4 below
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
                            for i in m:
                                j = ''.join(''.join(elems) for elems in i)
                                role = re.sub(r'' + firm + r'', '', j)
                                role = re.sub(r',', '', role)
                                if len(role) > 0:
                                    break

                            # Combine last name, person, firm, and role, for that person.
                            row = [last, person, firm, role, n]

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
                    m = re.findall(r'\b' + last + r'\b', stz4, re.IGNORECASE)
                    if len(m) > 0:
                        prow = [per[0], per[1], per[2], per[3], n]
                        prows.append(prow)

            # Cumulative list of people from people list referenced in sentences
            # Will refer to this in cases of pronouns (e.g., "she said"), etc.
            for per in people:
                y = re.findall(r'\b' + per[0] + r'\b', stz4, re.IGNORECASE)
                if len(y) > 0:
                    per[4] = n
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
                s = re.findall(r'' + stext + '', stz4, re.IGNORECASE)
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
                            x1 = re.sub(r'(' + qs + r')', '', x1)

                            if len(x1) > 0:
                                # if there is a said equivalent and quote but no person, check if pronoun near said equivalent
                                k1 = re.findall(
                                    r'(' + stext + r')\W+(?:\w+\W+){0,' + str(mxpron) + r'}?(' + pron + r')',
                                    stz4, re.IGNORECASE)
                                k2 = re.findall(
                                    r'(' + pron + r')\W+(?:\w+\W+){0,' + str(mxpron) + r'}?(' + stext + r')',
                                    stz4, re.IGNORECASE)
                                k = k1 + k2
                                k = ''.join(''.join(elems) for elems in k)
                                k = re.sub(r'(' + qs + r')', '', k)
                                if len(k) > 0:

                                    quote = x1
                                    last = sppl[-1][0]
                                    person = sppl[-1][1]
                                    firm = sppl[-1][2]
                                    role = sppl[-1][3]
                                    qtype = "single sentence quote"
                                    qsaid = "yes"
                                    qpref = "no"
                                    comment = "pronoun"
                                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid,
                                                             qpref,
                                                             comment,
                                                             AN, date, source, regions, subjects, misc, industries,
                                                             focfirm, by, comp, ment, stence]
                                    dx += 1
                                # If said equiv and quote, but no people or pronoun, check if organization listed, as in "BMW's spokesman said..."
                                elif len(k) == 0:
                                    if len(orgs) == 1:
                                        # check if spokeperson
                                        org = orgs[0]
                                        k3 = re.findall(r'' + org + r'.+(' + rep + r')', stz4, re.IGNORECASE)
                                        k4 = re.findall(r'' + rep + r'.+(' + org + r')', stz4, re.IGNORECASE)
                                        k = k3 + k4
                                        k = ''.join(''.join(elems) for elems in k)
                                        k = re.sub(r'(' + qs + r')', '', k)
                                        if len(k) > 0:
                                            last = "NA"
                                            # The person would be the representative's reference, removing the org name
                                            person = "NA"
                                            firm = org
                                            role = re.sub(r'' + org + '', '', k)
                                            quote = x1
                                            qtype = "single sentence quote"
                                            qsaid = "yes"
                                            qpref = "no"
                                            comment = "Org representative"
                                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype,
                                                                     qsaid,
                                                                     qpref,
                                                                     comment, AN, date, source, regions, subjects,
                                                                     misc, industries, focfirm, by, comp, ment,
                                                                     stence]
                                            dx += 1
                                        elif len(k) == 0:
                                            last = "NA"
                                            person = "NA"
                                            firm = org
                                            role = "NA"
                                            quote = x1
                                            qtype = "single sentence quote"
                                            qsaid = "yes"
                                            qpref = "no"
                                            comment = "from org, nondiscriptives"
                                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype,
                                                                     qsaid,
                                                                     qpref,
                                                                     comment, AN, date, source, regions, subjects,
                                                                     misc, industries, focfirm, by, comp, ment,
                                                                     stence]
                                            dx += 1
                                    elif len(orgs) > 1:
                                        last = "NA"
                                        person = "NA"
                                        firm = ''.join(''.join(elems) for elems in orgs)
                                        role = "NA"
                                        quote = x1
                                        qtype = "Information - non-discript single sentence quote"
                                        qsaid = "yes"
                                        qpref = "no"
                                        comment = "Review - from nondiscriptives"
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
                                                quote = x1
                                                last = sppl[-1][0]
                                                person = sppl[-1][1]
                                                firm = sppl[-1][2]
                                                role = sppl[-1][3]
                                                qtype = "single sentence quote"
                                                qsaid = "yes"
                                                qpref = "no"
                                                comment = "pronoun"
                                                df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype,
                                                                         qsaid,
                                                                         qpref, comment, AN, date, source, regions,
                                                                         subjects, misc, industries, focfirm, by,
                                                                         comp, ment, stence]
                                                dx += 1
                                            # if no person in previous sentence, then check if quote in previous sentence
                                            elif l > 1:
                                                # Check if there's something in DF:
                                                if dx>0:
                                                    s = df.loc[len(df.index) - 1].at['sent']
                                                    t = n - int(s)
                                                    if t == 1:
                                                        quote = x1
                                                        last = df.loc[len(df.index) - 1].at['last']
                                                        person = df.loc[len(df.index) - 1].at['person']
                                                        firm = df.loc[len(df.index) - 1].at['firm']
                                                        role = df.loc[len(df.index) - 1].at['role']
                                                        qtype = "single sentence quote"
                                                        qsaid = "yes"
                                                        qpref = "no"
                                                        comment = "pronoun - reference partial profile"
                                                        df.loc[len(df.index)] = [n, person, last, role, firm, quote,
                                                                                 qtype,
                                                                                 qsaid,
                                                                                 qpref, comment, AN, date, source,
                                                                                 regions, subjects, misc,
                                                                                 industries, focfirm, by, comp,
                                                                                 ment, stence]
                                                        dx += 1



                            # If said equivalent, no people and no quote, check if plural present
                            elif len(x1) == 0:
                                m = re.findall(r'(' + plurals + r')', stz4, re.IGNORECASE)

                                if len(m) > 0:
                                    last = "NA"
                                    person = "NA"
                                    firm = "NA"
                                    role = ''.join(''.join(elems) for elems in m)
                                    quote = stence
                                    qtype = "paraphrase"
                                    qsaid = "yes"
                                    qpref = "no"
                                    comment = "from plural individuals"
                                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid,
                                                             qpref,
                                                             comment,
                                                             AN, date, source, regions, subjects, misc, industries,
                                                             focfirm, by, comp, ment, stence]
                                    dx += 1


                                # If said equivalent, no people, no quote, no plural, check if pronoun, rep, and finally
                                # Org
                                elif len(m) == 0:

                                    k1 = re.findall(
                                        r'(' + stext + r')\W+(?:\w+\W+){0,' + str(mxpron) + r'}?(' + pron + r')',
                                        stz4, re.IGNORECASE)
                                    k2 = re.findall(
                                        r'(' + pron + r')\W+(?:\w+\W+){0,' + str(mxpron) + r'}?(' + stext + r')',
                                        stz4, re.IGNORECASE)
                                    k = k1 + k2
                                    k = ''.join(''.join(elems) for elems in k)
                                    k = re.sub(r'(' + qs + r')', '', k)
                                    if len(k) > 0:
                                        quote = stence
                                        last = sppl[-1][0]
                                        person = sppl[-1][1]
                                        firm = sppl[-1][2]
                                        role = sppl[-1][3]
                                        qtype = "paraphrase"
                                        qsaid = "yes"
                                        qpref = "no"
                                        comment = "pronoun"
                                        df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid,
                                                                 qpref,
                                                                 comment, AN, date, source, regions, subjects, misc,
                                                                 industries, focfirm, by, comp, ment, stence]
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
                                            p1a = re.findall(r'(' + l1 + r')', stence, re.IGNORECASE)
                                            p1b = re.findall(r'(' + l2 + r')', stence, re.IGNORECASE)
                                            p1c = re.findall(r'(' + l3 + r')', stence, re.IGNORECASE)
                                            p1d = re.findall(r'(' + l4 + r')', stence, re.IGNORECASE)
                                            p1e = re.findall(r'(' + l5 + r')', stence)
                                            p1 = p1a + p1b + p1c + p1d + p1e
                                            p1 = ''.join(''.join(elems) for elems in p1)
                                            if len(p1) > 0:
                                                quote = stence
                                                last = sppl[-1][0]
                                                person = sppl[-1][1]
                                                firm = sppl[-1][2]
                                                role = sppl[-1][3]
                                                qtype = "paraphrase"
                                                qsaid = "yes"
                                                qpref = "no"
                                                comment = "title reference"
                                                df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype,
                                                                         qsaid, qpref,
                                                                         comment, AN, date, source, regions,
                                                                         subjects, misc, industries, focfirm, by,
                                                                         comp, ment, stence]
                                                dx += 1
                                            elif len(p1) == 0:
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
                                                            last = "NA"
                                                            person = "NA"
                                                            firm = org
                                                            role = m
                                                            qtype = "paraphrase"
                                                            qsaid = "yes"
                                                            qpref = "one or more firms"
                                                            comment = "a firm representative said something"
                                                            quote = stence
                                                            df.loc[len(df.index)] = [n, person, last, role, firm,
                                                                                     quote,
                                                                                     qtype,
                                                                                     qsaid,
                                                                                     qpref, comment,
                                                                                     AN, date, source, regions,
                                                                                     subjects, misc, industries,
                                                                                     focfirm, by, comp, ment,
                                                                                     stence]
                                                            dx += 1
                                                        # If not a rep, see if just one company
                                                        # If so, the company said something
                                                        elif len(m) == 0:
                                                            if len(orgs) == 1:
                                                                quote = stence
                                                                last = "NA"
                                                                person = "NA"
                                                                firm = org
                                                                role = "NA"
                                                                qtype = "paraphrase"
                                                                qsaid = "yes"
                                                                qpref = "no"
                                                                comment = 'One firm said something'
                                                                df.loc[len(df.index)] = [n, person, last, role,
                                                                                         firm, quote,
                                                                                         qtype, qsaid, qpref,
                                                                                         comment, AN, date, source,
                                                                                         regions, subjects, misc,
                                                                                         industries, focfirm, by,
                                                                                         comp, ment, stence]
                                                                dx += 1
                                                            # If not, see if multiple companies
                                                            # If so, add line for each org and said, but flag to review
                                                            elif len(orgs) > 1:
                                                                for org in orgs:
                                                                    quote = stence
                                                                    last = "NA"
                                                                    person = "NA"
                                                                    firm = org
                                                                    role = "NA"
                                                                    qtype = "paraphrase"
                                                                    qsaid = "yes"
                                                                    qpref = "no"
                                                                    comment = "Review - paraphrase from multiple firms"
                                                                    df.loc[len(df.index)] = [n, person, last, role,
                                                                                             firm, quote,
                                                                                             qtype, qsaid, qpref,
                                                                                             comment, AN, date, source,
                                                                                             regions, subjects, misc,
                                                                                             industries, focfirm, by,
                                                                                             comp, ment, stence]
                                                                    dx += 1






                            # Because said equivalent but no full quote, check if it's a beginning of multi-sentence quote
                            # If so, assume it would have to be from the most recent referenced person
                            # As in "he said", referring to the most recent person
                            else:
                                # check if a quote begins in this sentence:
                                q1 = re.findall(r'(' + start + r').+', stence, re.IGNORECASE)
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
                                    last = sppl[-1][0]
                                    person = sppl[-1][1]
                                    firm = sppl[-1][2]
                                    role = sppl[-1][3]
                                    qtype = "multi-sentence quote"
                                    qsaid = "yes"
                                    qpref = "no"
                                    comment = "high accuracy"
                                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid,
                                                             qpref,
                                                             comment,
                                                             AN, date, source, regions, subjects, misc, industries,
                                                             focfirm, by, comp, ment, stence]
                                    dx += 1


                    # if said something and exactly one person from list, do stuff
                    elif len(prows) == 1:
                        ##################################
                        # ONE AND ONLY ONE PEOPLE REFERENCE
                        ##################################

                        last = prows[0][0]

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
                            quote = re.sub(r'' + stext + r'\Z', '', quote)
                            last = prows[0][0]
                            person = prows[0][1]
                            firm = prows[0][2]
                            role = prows[0][3]
                            qtype = "single sentence quote"
                            qsaid = "yes"
                            qpref = "single"
                            comment = "high accuracy"
                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref,
                                                     comment,
                                                     AN, date,
                                                     source, regions, subjects, misc, industries, focfirm, by, comp,
                                                     ment, stence]
                            dx += 1


                        # Check if the multiple sentence quote
                        elif len(s1) == 0:
                            # Because said equivalent but no full quote, check if it's a beginning of multi-sentence quote
                            # If so, assume it would have to be from the most recent referenced person
                            # check if a quote begins in this sentence:
                            q1 = re.findall(r'(' + last + r').+(' + stext + r').+(' + start + r')(.+)', stence,
                                            re.IGNORECASE)
                            q1 = ''.join(''.join(elems) for elems in q1)
                            q1 = re.sub(r'(' + qs + r')', '', q1)
                            q1 = re.sub(r'' + last + r'', '', q1)
                            q1 = re.sub(r'(' + stext + r')', '', q1)

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
                                last = prows[0][0]
                                person = prows[0][1]
                                firm = prows[0][2]
                                role = prows[0][3]
                                qtype = "multi-sentence quote"
                                qsaid = "yes"
                                qpref = "single"
                                comment = "high accuracy"
                                df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref,
                                                         comment, AN,
                                                         date, source, regions, subjects, misc, industries, focfirm,
                                                         by, comp, ment, stence]
                                dx += 1
                            # Check if the name is actually inside the quote
                            # If so, the speaker should be the previous referenced person
                            elif len(q1) == 0:
                                s3 = re.findall(r'(' + start + r')(.+' + last + r'.+)(' + end + r')', stence,
                                                re.IGNORECASE)
                                s3 = ''.join(''.join(elems) for elems in s3)
                                s3 = re.sub(r'(' + qs + r')', '', s3)

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
                                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid,
                                                             qpref,
                                                             comment,
                                                             AN, date, source, regions, subjects, misc, industries,
                                                             focfirm, by, comp, ment, stence]
                                    dx += 1

                                elif len(s3) == 0:
                                    q1a = re.findall(r'(' + end + r').+(' + stext + r').+(' + last + r')', stence,
                                                     re.IGNORECASE)
                                    q1b = re.findall(r'(' + end + r').+(' + last + r').+(' + stext + r')', stence,
                                                     re.IGNORECASE)
                                    q1 = q1a + q1b
                                    q1 = ''.join(''.join(elems) for elems in q1)
                                    q = re.sub(r'(' + qs + r')', '', q1)
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
                                        df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid,
                                                                 qpref,
                                                                 comment, AN, date, source, regions, subjects, misc,
                                                                 industries, focfirm, by, comp, ment, stence]
                                        dx += 1


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
                            s1a = re.findall(
                                r'(' + start + r')(.+)(' + end + r')(?=.+(' + last + r')\W+(?:\w+\W+){0,' + str(
                                    mxlds) + r'}?' + stext + r')', stence, re.IGNORECASE)
                            # Said equiv then last name then a quote (distance not matter here because sometimes full title in between)
                            s1b = re.findall(
                                r'(' + last + r').+(' + stext + 'r).+(' + start + r')(.+)(' + end + r')',
                                stence,
                                re.IGNORECASE)
                            s1 = s1a + s1b
                            s1 = ''.join(''.join(elems) for elems in s1)
                            s1 = re.sub(r'(' + qs + r')', '', s1)
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
                                df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref,
                                                         comment, AN,
                                                         date, source, regions, subjects, misc, industries, focfirm,
                                                         by, comp, ment, stence]
                                dx += 1
                            # check if a quote begins in this sentence:
                            elif len(s1) == 0:
                                q1a = re.findall(r'(' + last + r').+(' + stext + r').+(' + start + r')(.+)', stence,
                                                 re.IGNORECASE)
                                # said equiv then last name then a quote
                                q1b = re.findall(r'(' + stext + r').+(' + last + r').+(' + start + r')(.+)', stence,
                                                 re.IGNORECASE)
                                q1 = q1a + q1b
                                q1 = ''.join(''.join(elems) for elems in q1)
                                q1 = re.sub(r'(' + qs + r')', '', q1)
                                q1 = re.sub(r'(' + stext + r')', '', q1)
                                q1 = re.sub(r'' + last + r'', '', q1)

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
                                    quote = re.sub(r'' + stext + r'\Z', '', quote)
                                    qtype = "multi-sentence quote"
                                    qsaid = "yes"
                                    qpref = "multiple"
                                    comment = "review"
                                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid,
                                                             qpref,
                                                             comment,
                                                             AN, date, source, regions, subjects, misc, industries,
                                                             focfirm, by, comp, ment, stence]
                                    dx += 1
                                elif len(q1) == 0:
                                    quote = stence
                                    qtype = "paraphrase"
                                    qsaid = "yes"
                                    qpref = "multiple"
                                    comment = "review"
                                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid,
                                                             qpref,
                                                             comment,
                                                             AN, date, source, regions, subjects, misc, industries,
                                                             focfirm, by, comp, ment, stence]
                                    dx += 1


                elif len(s) == 0:
                ##################################
                # SENTENCES WITH NO "SAID" EQUIVALENT
                ##################################
                    # Check if no person:
                    if len(prows) == 0:
                        ##################################
                        # NO PEOPLE REFERENCE
                        ##################################
                        # Since no clear person reference, check if 2+ words in a quote:
                        x1 = re.findall(r'(' + start + r')(\b.+\b\s+\b.+\b.*)(' + end + r')', stence, re.IGNORECASE)
                        x1 = ''.join(''.join(elems) for elems in x1)
                        x1 = re.sub(r'(' + qs + r')', '', x1)
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
                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref,
                                                     comment,
                                                     AN, date,
                                                     source, regions, subjects, misc, industries, focfirm, by, comp,
                                                     ment, stence]
                            dx += 1
                        # Because no said equiv or full quote, check if it's a beginning of multi-sentence quote
                        elif len(x1) == 0:
                            # Check if just one word quoted, keep going and do nothing
                            g = re.findall(r'(' + start + r')(\b.+\b)(' + end + r')', stence, re.IGNORECASE)
                            g = ''.join(''.join(elems) for elems in g)
                            if len(g) > 1:
                                pass
                            # If not, check if a quote begins in this sentence:
                            else:
                                q1 = re.findall(r'(' + start + r')(.+)', stence, re.IGNORECASE)
                                q1 = ''.join(''.join(elems) for elems in q1)

                                # If a quote does begin in this sentence, concatinate the subsequent sentences until the quote ends
                                if len(q1) > 0:
                                    # set q2a variable as blank and then update if multi-sentence
                                    q2a = ""
                                    for s in range(n, len(doc.sentences)):
                                        j = doc.sentences[s]
                                        t = j.text
                                        q2 = re.findall(r'(.+)(' + end + r')', t, re.IGNORECASE)
                                        q2 = ''.join(''.join(elems) for elems in q2)
                                        if len(q2) == 0:
                                            q1 = q1 + " " + t
                                        else:
                                            q1 = q1 + " " + q2
                                            q2a = re.findall(r'(' + end + r')(.+)', t, re.IGNORECASE)
                                            q2a = ''.join(''.join(elems) for elems in q2a)
                                            break
                                    # If multi-sentence
                                    if len(q1) > 0:
                                        z1 = re.findall(r'(' + l1 + r')', q2a, re.IGNORECASE)
                                        z2 = re.findall(r'(' + l2 + r')', q2a, re.IGNORECASE)
                                        z3 = re.findall(r'(' + l3 + r')', q2a, re.IGNORECASE)
                                        z4 = re.findall(r'(' + l4 + r')', q2a, re.IGNORECASE)
                                        z5 = re.findall(r'(' + l5 + r')', q2a)
                                        m = z1 + z2 + z3 + z4 + z5

                                        z = ''.join(''.join(elems) for elems in m)

                                        q1 = re.sub(r'(' + qs + r')', '', q1)

                                        if len(z) > 0:
                                            quote = q1
                                            last = "NA"
                                            person = "NA"
                                            firm = "NA"
                                            role = ""
                                            for i in m:
                                                j = ''.join(''.join(elems) for elems in i)
                                                role = re.sub(r',', '', j)
                                                if len(role) > 0:
                                                    break
                                            qtype = "multi-sentence quote"
                                            qsaid = "no"
                                            qpref = "no"
                                            comment = "high accuracy"
                                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype,
                                                                     qsaid,
                                                                     qpref,
                                                                     comment, AN,
                                                                     date, source, regions, subjects, misc,
                                                                     industries, focfirm, by, comp, ment, stence]
                                            dx += 1

                                        elif len(z) == 0:
                                            quote = q1
                                            last = sppl[-1][0]
                                            person = sppl[-1][1]
                                            firm = sppl[-1][2]
                                            role = sppl[-1][3]
                                            qtype = "multi-sentence quote"
                                            qsaid = "no"
                                            qpref = "no"
                                            comment = "high accuracy"
                                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype,
                                                                     qsaid,
                                                                     qpref, comment, AN,
                                                                     date, source, regions, subjects, misc,
                                                                     industries, focfirm, by, comp, ment, stence]
                                            dx += 1

                    ##################################
                    # ONE OR MORE PEOPLE REFERENCED
                    ##################################
                    elif len(prows) > 0:
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
                                quote = re.sub(r'' + stext + r'\Z', '', quote)
                                last = "NA"
                                person = "NA"
                                firm = "NA"
                                role = "NA"
                                qtype = "single sentence quote"
                                qsaid = "no"
                                qpref = "one or more"
                                comment = "Review, whether quote or something else, not clear"
                                df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref,
                                                         comment, AN,
                                                         date,
                                                         source, regions, subjects, misc, industries, focfirm, by,
                                                         comp, ment, stence]
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
                                quote = re.sub(r'' + stext + r'\Z', '', quote)
                                last = sppl[-1][0]
                                person = sppl[-1][1]
                                firm = sppl[-1][2]
                                role = sppl[-1][3]
                                qtype = "multi-sentence quote"
                                qsaid = "no"
                                qpref = "one or more"
                                comment = "high accuracy - prior person referenced"
                                df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref,
                                                         comment, AN,
                                                         date, source, regions, subjects, misc, industries, focfirm,
                                                         by, comp, ment, stence]
                                dx += 1
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
                                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid,
                                                             qpref,
                                                             comment,
                                                             AN, date, source, regions, subjects, misc, industries,
                                                             focfirm, by, comp, ment, stence]
                                    dx += 1

            ##################################
            # NO PEOPLE IN CUMULATIVE LIST YET
            ##################################
            if len(people) == 0:

                # This should be rare because typically the first mentioning of a person includes the title and organization.
                # Because no person (i.e., with firm, title, and name) associated with it yet, see if at least a Person:
                # If exactly 1 person, check if quotes. No need to check for pronouns yet because no reference yet:

                if len(ppl) == 1:
                    per = ppl[0]
                    # Quote then person then said equivalent
                    s1a = re.findall(r'(' + start + r')(.+)(' + end + r').+(' + per + r').+(' + stext + r')',
                                     stence,
                                     re.IGNORECASE)
                    s1b = re.findall(r'(' + start + r')(.+)(' + end + r').+(' + stext + r').+(' + per + r')',
                                     stence,
                                     re.IGNORECASE)
                    # Said equiv then last name then a quote (distance not matter here because sometimes full title in between)
                    s1c = re.findall(r'(' + per + r').+(' + stext + r').+(' + start + r')(.+)(' + end + r')',
                                     stence,
                                     re.IGNORECASE)
                    s1d = re.findall(r'(' + stext + r').+(' + per + r').+(' + start + r')(.+)(' + end + r')',
                                     stence,
                                     re.IGNORECASE)
                    s1 = s1a + s1b + s1c + s1d
                    s1 = ''.join(''.join(elems) for elems in s1)
                    s1 = re.sub(r'(' + qs + r')', '', s1)
                    s1 = re.sub(r'(' + stext + r')', '', s1)
                    s1 = re.sub(r'' + per + r'', '', s1)
                    if len(s1) > 0:
                        # If so, check if person is also near title
                        o7a = re.findall(
                            r'\b(?:(' + per + r')\W+(?:\w+\W+){0,' + str(
                                mxorg1) + r'}?(' + l1 + r'))',
                            stz4, re.IGNORECASE)
                        o7b = re.findall(
                            r'\b(?:(' + per + r')\W+(?:\w+\W+){0,' + str(
                                mxorg1) + r'}?(' + l2 + r'))',
                            stz4, re.IGNORECASE)
                        o7c = re.findall(
                            r'\b(?:(' + per + r')\W+(?:\w+\W+){0,' + str(
                                mxorg1) + r'}?(' + l3 + r'))',
                            stz4, re.IGNORECASE)
                        o7d = re.findall(
                            r'\b(?:(' + per + r')\W+(?:\w+\W+){0,' + str(
                                mxorg1) + r'}?(' + l4 + r'))',
                            stz4, re.IGNORECASE)
                        o7e = re.findall(
                            r'\b(?:(' + per + r')\W+(?:\w+\W+){0,' + str(
                                mxorg1) + r'}?(' + l5 + r'))',
                            stz4)
                        o8a = re.findall(
                            r'\b(?:(' + l1 + r')\W+(?:\w+\W+){0,' + str(
                                mxorg1) + r'}?(' + per + r'))',
                            stz4, re.IGNORECASE)
                        o8b = re.findall(
                            r'\b(?:(' + l2 + r')\W+(?:\w+\W+){0,' + str(
                                mxorg1) + r'}?(' + per + r'))',
                            stz4, re.IGNORECASE)
                        o8c = re.findall(
                            r'\b(?:(' + l3 + r')\W+(?:\w+\W+){0,' + str(
                                mxorg1) + r'}?(' + per + r'))',
                            stz4, re.IGNORECASE)
                        o8d = re.findall(
                            r'\b(?:(' + l4 + r')\W+(?:\w+\W+){0,' + str(
                                mxorg1) + r'}?(' + per + r'))',
                            stz4, re.IGNORECASE)
                        o8e = re.findall(
                            r'\b(?:(' + l5 + r')\W+(?:\w+\W+){0,' + str(
                                mxorg1) + r'}?(' + per + r'))',
                            stz4)

                        m = o7a + o7b + o7c + o7d + o7e + o8a + o8b + o8c + o8d + o8e

                        ol = ''.join(''.join(elems) for elems in m)
                        # If name and person are close to title, assign
                        if len(ol) > 0:

                            role = ""
                            for i in m:
                                j = ''.join(''.join(elems) for elems in i)
                                role = re.sub(r'' + per + r'', '', j)
                                role = re.sub(r',', '', role)
                                if len(role) > 0:
                                    break
                            if len(corgs) == 1:
                                firm = corgs[0]
                                comment = "No full profile of person yet"
                            elif len(corgs) > 0:
                                firm = "NA"
                                comment = "review - no full profile of person yet"
                            person = per
                            sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                            if len(sfx) > 0:
                                last = person.split()[-2]
                            elif len(sfx) == 0:
                                last = person.split()[-1]

                            qtype = "single sentence quote"
                            qsaid = "yes"
                            qpref = "yes"
                            quote = s1
                            quote = re.sub(r'' + per + '', '', quote)
                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref,
                                                     comment, AN, date, source, regions, subjects, misc, industries,
                                                     focfirm, by, comp, ment, stence]
                            dx += 1
                        # If not title, then we mark the person and the quote
                        elif (ol) == 0:

                            person = per
                            sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                            if len(sfx) > 0:
                                last = person.split()[-2]
                            elif len(sfx) == 0:
                                last = person.split()[-1]
                            firm = "NA"
                            role = "NA"
                            qtype = "single sentence quote"
                            qsaid = "yes"
                            qpref = "yes"
                            comment = "review - no full profile of person yet"
                            quote = s1
                            quote = re.sub(r'' + per + '', '', quote)
                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref,
                                                     comment, AN,
                                                     date,
                                                     source, regions, subjects, misc, industries, focfirm, by, comp,
                                                     ment, stence]
                            dx += 1
                    elif len(s1) == 0:
                        # check if a quote begins in this sentence:
                        q1a = re.findall(r'(' + per + r').+(' + stext + r').+(' + start + r')(.+)', stence,
                                         re.IGNORECASE)
                        q1b = re.findall(r'(' + stext + r').+(' + per + r').+(' + start + r')(.+)', stence,
                                         re.IGNORECASE)
                        q1 = q1a + q1b
                        q1 = ''.join(''.join(elems) for elems in q1)
                        q1 = re.sub(r'(' + qs + r')', '', q1)
                        q1 = re.sub(r'(' + stext + r')', '', q1)
                        q1 = re.sub(r'' + per + r'', '', q1)

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

                            sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                            person = per
                            if len(sfx) > 0:
                                last = person.split()[-2]
                            elif len(sfx) == 0:
                                last = person.split()[-1]

                            firm = "NA"
                            role = "NA"
                            qtype = "multi-sentence quote"
                            qsaid = "yes"
                            qpref = "yes"
                            quote = q1
                            quote = re.sub(r'^' + stext + r'.+', '', quote)
                            quote = re.sub(r'' + stext + r'\Z', '', quote)
                            comment = "review - no full profile of person yet"
                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref,
                                                     comment,
                                                     AN, date,
                                                     source, regions, subjects, misc, industries, focfirm, by, comp,
                                                     ment, stence]
                            dx += 1


                        # If one person but no quote single or multiple, check for orgs
                        elif len(q1) == 0:
                            # If one org and one person, check if said:
                            if len(orgs) == 1:
                                # If so, then assign
                                s = re.findall(r'' + stext + '', stence, re.IGNORECASE)
                                if len(s) > 0:
                                    person = per
                                    sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                                    if len(sfx) > 0:
                                        last = person.split()[-2]
                                    elif len(sfx) == 0:
                                        last = person.split()[-1]

                                    firm = orgs[0]
                                    role = "NA"
                                    qtype = "paraphrase"
                                    qsaid = "yes"
                                    qpref = "one firm"
                                    comment = "statement by the firm"
                                    quote = stence
                                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid,
                                                             qpref,
                                                             comment, AN, date, source, regions, subjects, misc,
                                                             industries, focfirm, by, comp, ment, stence]
                                    dx += 1
                                elif len(s) == 0:
                                    person = per
                                    sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                                    if len(sfx) > 0:
                                        last = person.split()[-2]
                                    elif len(sfx) == 0:
                                        last = person.split()[-1]

                                    firm = orgs[0]
                                    role = "NA"
                                    qtype = "information"
                                    qsaid = "no"
                                    qpref = "one firm"
                                    comment = "information about individual"
                                    quote = stence
                                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid,
                                                             qpref,
                                                             comment,
                                                             AN,
                                                             date, source, regions, subjects, misc, industries,
                                                             focfirm, by, comp, ment, stence]
                                    dx += 1
                            # if 2+ orgs, check for said
                            elif len(orgs) > 1:
                                s = re.findall(r'' + stext + '', stence, re.IGNORECASE)
                                if len(s) > 0:
                                    person = per
                                    sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                                    if len(sfx) > 0:
                                        last = person.split()[-2]
                                    elif len(sfx) == 0:
                                        last = person.split()[-1]

                                    firm = ''.join(''.join(elems) for elems in orgs)
                                    role = "NA"
                                    qtype = "paraphrase"
                                    qsaid = "yes"
                                    qpref = "multiple firms"
                                    comment = "review - statement by one of multiple firms"
                                    quote = stence
                                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid,
                                                             qpref,
                                                             comment, AN,
                                                             date, source, regions, subjects, misc, industries,
                                                             focfirm, by, comp, ment, stence]
                                    dx += 1

                                elif len(s) == 0:
                                    person = per
                                    sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                                    if len(sfx) > 0:
                                        last = person.split()[-2]
                                    elif len(sfx) == 0:
                                        last = person.split()[-1]

                                    firm = ''.join(''.join(elems) for elems in orgs)
                                    role = "NA"
                                    qtype = "information"
                                    qsaid = "yes"
                                    qpref = "multiple firms"
                                    comment = "one person and multiple firms"
                                    quote = stence
                                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid,
                                                             qpref,
                                                             comment, AN,
                                                             date, source, regions, subjects, misc, industries,
                                                             focfirm, by, comp, ment, stence]
                                    dx += 1


                            elif len(orgs) == 0:
                                s = re.findall(r'' + stext + '', stence, re.IGNORECASE)
                                if len(s) > 0:
                                    # If no quote single or multiple, no org, and no people list, then information
                                    person = per
                                    sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                                    if len(sfx) > 0:
                                        last = person.split()[-2]
                                    elif len(sfx) == 0:
                                        last = person.split()[-1]

                                    firm = "NA"
                                    role = "NA"
                                    qtype = "paraphrase"
                                    qsaid = "yes"
                                    qpref = "yes"
                                    quote = stence
                                    comment = "From one person"
                                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid,
                                                             qpref,
                                                             comment, AN,
                                                             date, source, regions, subjects, misc, industries,
                                                             focfirm, by, comp, ment, stence]
                                    dx += 1
                                if len(s) == 0:
                                    # If no quote single or multiple, no org, and no people list, then information
                                    person = per
                                    sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                                    if len(sfx) > 0:
                                        last = person.split()[-2]
                                    elif len(sfx) == 0:
                                        last = person.split()[-1]

                                    firm = "NA"
                                    role = "NA"
                                    qtype = "information"
                                    qsaid = "yes"
                                    qpref = "yes"
                                    quote = stence
                                    comment = "Information about one person"
                                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid,
                                                             qpref,
                                                             comment, AN,
                                                             date, source, regions, subjects, misc, industries,
                                                             focfirm, by, comp, ment, stence]
                                    dx += 1

                elif len(ppl) > 1:
                    # Very rare. If more than one person (without titles/orgs), include any organizations in the sentence
                    # flag for review
                    for per in ppl:
                        quote = stence
                        person = per
                        sfx = re.findall(r'(' + suf + r')', person, re.IGNORECASE)
                        if len(sfx) > 0:
                            last = person.split()[-2]
                        elif len(sfx) == 0:
                            last = person.split()[-1]

                        firm = ''.join(''.join(elems) for elems in orgs)
                        role = "NA"
                        qtype = "Information"
                        qsaid = "no"
                        qpref = "more than one"
                        comment = "review - no full profile of person yet"
                        df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref, comment,
                                                 AN,
                                                 date,
                                                 source, regions, subjects, misc, industries, focfirm, by, comp,
                                                 ment, stence]
                        dx += 1
                elif len(ppl) == 0:

                    if len(orgs) == 1:
                        s = re.findall(r'' + stext + '', stence, re.IGNORECASE)
                        if len(s) > 0:
                            last = "NA"
                            person = "NA"
                            firm = orgs[0]
                            role = "NA"
                            qtype = "paraphrase"
                            qsaid = "yes"
                            qpref = "one - firm"
                            comment = "statement by the firm"
                            quote = stence
                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref,
                                                     comment,
                                                     AN, date,
                                                     source, regions, subjects, misc, industries, focfirm, by, comp,
                                                     ment, stence]
                            dx += 1
                        elif len(s) == 1:
                            last = "NA"
                            person = "NA"
                            firm = orgs[0]
                            role = "NA"
                            qtype = "information"
                            qsaid = "yes"
                            qpref = "one - firm"
                            comment = "info about a firm"
                            quote = stence
                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid, qpref,
                                                     comment,
                                                     AN, date,
                                                     source, regions, subjects, misc, industries, focfirm, by, comp,
                                                     ment, stence]
                            dx += 1

                    elif len(orgs) > 1:
                        s = re.findall(r'' + stext + '', stence, re.IGNORECASE)

                        if len(s) > 0:
                            # If someone said something, see if it's a representative of a firm
                            for org in orgs:
                                m1 = re.findall(r'' + org + r'\s+(' + rep + r')', stz4, re.IGNORECASE)
                                m2 = re.findall(r'(' + rep + r')\s+of\s+' + org + r'', stz4, re.IGNORECASE)
                                m = m1 + m2
                                m = ''.join(''.join(elems) for elems in m)
                                m = re.sub(r'(' + org + r')', '', m)
                                if len(m) > 0:
                                    last = "NA"
                                    person = "NA"
                                    firm = org
                                    role = m
                                    qtype = "paraphrase"
                                    qsaid = "yes"
                                    qpref = "one or more firms"
                                    comment = "one or more firm representatives said something"
                                    quote = stence
                                    df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid,
                                                             qpref,
                                                             comment,
                                                             AN, date, source, regions, subjects, misc, industries,
                                                             focfirm, by, comp, ment, stence]
                                    dx += 1


                    # No people list, no people, no org, see if a quote
                    elif len(orgs) == 0:
                        # see if quote
                        # Quote then person then said equivalent
                        s1a = re.findall(r'(' + start + r')(.+)(' + end + r').+(' + stext + r')', stence,
                                         re.IGNORECASE)
                        s2a = re.findall(r'(' + start + r')(.+)(' + end + r').+(' + stext + r')', stence,
                                         re.IGNORECASE)
                        s = s1a + s2a
                        s = ''.join(''.join(elems) for elems in s)
                        s = re.sub(r'(' + qs + r')', '', s)
                        s = re.sub(r'(' + stext + r')', '', s)
                        if len(s) > 0:
                            # Check for pronoun:
                            s3a = re.findall(r'' + pron + r'', stz4, re.IGNORECASE)
                            if len(s3a) > 0:
                                # Check if there's something in sppl:
                                if len(sppl) > 0:
                                    # check when the last person reference was
                                    # The second items, the item in the person list, must be last because
                                    # sppl has n ongoing
                                    m = int(sppl[-1][-1])
                                    l = n - int(m)
                                    # If person in the previous sentence, use that one
                                    if l == 1:
                                        quote = s
                                        last = sppl[-1][0]
                                        person = sppl[-1][1]
                                        firm = sppl[-1][2]
                                        role = sppl[-1][3]
                                        qtype = "single sentence quote"
                                        qsaid = "yes"
                                        qpref = "no"
                                        comment = "pronoun - from previous sentence"
                                        df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid,
                                                                 qpref, comment, AN, date, source, regions,
                                                                 subjects, misc, industries, focfirm, by, comp,
                                                                 ment, stence]
                                        dx += 1
                                    # if no person in previous sentence, then check if quote in previous sentence
                                    # You should attribute this quote to the person attributed to the previous sentence.
                                    elif l > 1:
                                        # Check if there's something in DF:
                                        if dx>0:
                                            s = df.loc[len(df.index) - 1].at['sent']
                                            t = n - int(s)
                                            if t == 1:
                                                quote = s
                                                last = df.loc[len(df.index) - 1].at['last']
                                                person = df.loc[len(df.index) - 1].at['person']
                                                firm = df.loc[len(df.index) - 1].at['firm']
                                                role = df.loc[len(df.index) - 1].at['role']
                                                qtype = "single sentence quote"
                                                qsaid = "yes"
                                                qpref = "no"
                                                comment = "pronoun - reference partial profile"
                                                df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype,
                                                                         qsaid,
                                                                         qpref, comment, AN, date, source, regions,
                                                                         subjects, misc, industries, focfirm, by,
                                                                         comp, ment, stence]
                                                dx += 1
                        # if no quote, see if beginning of multi-sentence quote:
                        elif len(s) == 0:
                            # check if a quote begins in this sentence:
                            q1a = re.findall(r'^\s*(' + start + r')(.+)', stence, re.IGNORECASE)
                            q1b = re.findall(r'(' + stext + r').+\s+(' + start + r')(.+)', stence, re.IGNORECASE)

                            q1 = q1a + q1b
                            q1 = ''.join(''.join(elems) for elems in q1)
                            q1 = re.sub(r'(' + qs + r')', '', q1)
                            q1 = re.sub(r'(' + stext + r')', '', q1)

                            # If a quote does begin in this sentence, concatinate the subsequent sentences until the quote ends
                            if len(q1) > 0:
                                q2a = ""
                                for s in range(n, len(doc.sentences)):
                                    j = doc.sentences[s]
                                    t = j.text
                                    q2 = re.findall(r'(.+)(' + end + r')', t, re.IGNORECASE)
                                    q2 = ''.join(''.join(elems) for elems in q2)
                                    q2 = re.sub(r'(' + qs + r')', '', q2)
                                    if len(q2) == 0:
                                        q1 = q1 + " " + t
                                    else:
                                        q1 = q1 + " " + q2
                                        q2a = re.findall(r'(' + end + r')(.+)', t, re.IGNORECASE)
                                        q2a = ''.join(''.join(elems) for elems in q2a)
                                        break
                                q1 = re.sub(r'(' + qs + r')', '', q1)
                                z1 = re.findall(r'(' + l1 + r')', q2a, re.IGNORECASE)
                                z2 = re.findall(r'(' + l2 + r')', q2a, re.IGNORECASE)
                                z3 = re.findall(r'(' + l3 + r')', q2a, re.IGNORECASE)
                                z4 = re.findall(r'(' + l4 + r')', q2a, re.IGNORECASE)
                                z5 = re.findall(r'(' + l5 + r')', q2a)
                                m = z1 + z2 + z3 + z4 + z5

                                z = ''.join(''.join(elems) for elems in m)
                                if len(z) > 0:
                                    # was information tied to previous sentence.
                                    # You should attribute this quote to the person (a partial profile)
                                    # attributed to the previous sentence.
                                    # Check if there's something in DF:
                                    if dx>0:
                                        s = df.loc[len(df.index) - 1].at['sent']
                                        t = n - int(s)
                                        if t == 1:
                                            quote = q1
                                            last = df.loc[len(df.index) - 1].at['last']
                                            person = df.loc[len(df.index) - 1].at['person']
                                            firm = df.loc[len(df.index) - 1].at['firm']
                                            role = ""
                                            for i in m:
                                                j = ''.join(''.join(elems) for elems in i)
                                                role = re.sub(r',', '', j)
                                                if len(role) > 0:
                                                    break
                                            qtype = "multi-sentence quote"
                                            qsaid = "maybe"
                                            qpref = "no"
                                            comment = "Title of previous partial profile person"
                                            df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype,
                                                                     qsaid,
                                                                     qpref, comment, AN, date, source, regions,
                                                                     subjects, misc, industries, focfirm, by, comp,
                                                                     ment, stence]
                                            dx += 1
                                # If not, check for pronouns in last sentence
                                elif len(z) == 0:
                                    s4a = re.findall(r'' + pron + r'.+' + stext + r'', q2a, re.IGNORECASE)
                                    s4b = re.findall(r'' + stext + r'.+' + pron + r'', q2a, re.IGNORECASE)
                                    z = s4a + s4b
                                    z = ''.join(''.join(elems) for elems in z)

                                    if len(z) > 0:
                                        # was information tied to previous sentence.
                                        # You should attribute this quote to the person (a partial profile)
                                        # attributed to the previous sentence.
                                        # Check if there's something in DF:
                                        if dx>0:
                                            s = df.loc[len(df.index) - 1].at['sent']
                                            t = n - int(s)
                                            if t == 1:
                                                quote = q1
                                                last = df.loc[len(df.index) - 1].at['last']
                                                person = df.loc[len(df.index) - 1].at['person']
                                                firm = df.loc[len(df.index) - 1].at['firm']
                                                role = df.loc[len(df.index) - 1].at['role']
                                                qtype = "multi-sentence quote"
                                                qsaid = "maybe"
                                                qpref = "no"
                                                comment = "pronoun of previous partial person"
                                                df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype,
                                                                         qsaid,
                                                                         qpref, comment, AN, date, source, regions,
                                                                         subjects, misc, industries, focfirm, by,
                                                                         comp, ment, stence]
                                                dx += 1
                            # If no quote single/multiple, see if any information
                            if len(q1) == 0:
                                if dx>0:
                                    y = df.loc[len(df.index) - 1].at['last']
                                    x = re.findall(r'(' + y + r')', stz4, re.IGNORECASE)
                                    if len(x) > 0:
                                        last = df.loc[len(df.index) - 1].at['last']
                                        person = df.loc[len(df.index) - 1].at['person']
                                        firm = df.loc[len(df.index) - 1].at['firm']
                                        role = df.loc[len(df.index) - 1].at['role']
                                        qtype = "information"
                                        qsaid = "no"
                                        qpref = "yes"
                                        comment = "review potentially info about an individual"
                                        quote = stence
                                        df.loc[len(df.index)] = [n, person, last, role, firm, quote, qtype, qsaid,
                                                                 qpref,
                                                                 comment, AN, date,
                                                                 source, regions, subjects, misc, industries,
                                                                 focfirm, by, comp, ment, stence]
                                        dx += 1

            n += 1



    #except:
     #   df.loc[len(df.index)] = [n, "Issue", "Issue", "Issue", "Issue", "Issue", "Issue", "Issue", "Issue",
     #                            "Issue", AN, date, source, "Issue", "Issue", "Issue", "Issue", "Issue", "Issue",
     #                            "Issue", "Issue", "Issue"]

    t2 = time.time()
    total = t1 - t0
    print(art + 1, AN, "time run:", round(t2 - t1, 2), "total hours:", round((t2 - t0) / (60 * 60), 2),
          "mean rate:", round((t2 - t0) / (art + 1), 2))

# output the chunk
# OUTFILE = f"{OUTDIR}/{INFILE.split('.')[0]}_df_quotes_{STARTROW.zfill(6)}_{ENDROW.zfill(6)}.xlsx"
# df.to_excel(OUTFILE)


df
# df.to_csv(r'C:\Users\danwilde\Dropbox (Penn)\Dissertation\Factiva\data1.csv')

# print(people)


# TO DO:
# Former leaders