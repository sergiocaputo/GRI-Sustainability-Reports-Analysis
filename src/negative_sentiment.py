import json
import os
import re
from GRIs import GRI

def count_negative_contexts(GRI_dir):
    neg_contexts = 0
    for doc in os.listdir(GRI_dir):
        doc = os.path.join(GRI_dir, doc)

        with open(doc) as file:
            json_GRI =  json.load(file)

        for k in GRI.keys():
            if json_GRI[k] != False and json_GRI[k]["context_sentiment"] == "neg":
                neg_contexts += 1
    return neg_contexts


def count_occurences(words, GRI_dir):
    occurences = []
    for w in words:
        occ = 0
        for doc in os.listdir(GRI_dir):
            doc = os.path.join(GRI_dir, doc)

            with open(doc) as file:
                json_GRI = json.load(file)

            for k in GRI.keys():
                if json_GRI[k] != False and json_GRI[k]["context_sentiment"] == "neg":
                    context = json_GRI[k]['context_before']+ " " + json_GRI[k]['context_after']
                    regex = r"\b" + w +"\w*"
                    
                    if len(re.findall(regex, context)) != 0:
                        occ += 1

        occurences.append(occ)
    return occurences


if __name__ == "__main__":

    neg_sentiment_words = ['rischi', 'corruzione', 'rifiuti', 'dannos', 'discriminator', 'malatti', 'inquinant', 'non conforme', 'spesa', ' emission',
                            'infortun', 'sanzion', 'incident', 'decess', 'pericolos' ,'violazion', 'mort']
                            
    neg_contexts_extracted = count_negative_contexts('GRI_reports_extracted')
    neg_contexts_ocr = count_negative_contexts('GRI_reports_ocr')

    neg_occ_ocr = count_occurences(neg_sentiment_words, 'GRI_reports_ocr')
    print("Negative contexts ocr: ", neg_contexts_ocr)
    print("Occurrences per neg root words in ocr: ", neg_occ_ocr)
    print("Sum occurrences neg root words: ", sum(neg_occ_ocr))
    print("Percentage per neg root words in neg contexts ocr: ", [round(o/neg_contexts_ocr, 2) for o in neg_occ_ocr])
    print()

    neg_occ_extracted = count_occurences(neg_sentiment_words, 'GRI_reports_extracted')
    print("Negative contexts extracted: ", neg_contexts_extracted)
    print("Occurrences per neg root words in extracted: ", neg_occ_extracted)
    print("Sum occurrences neg root words: ", sum(neg_occ_extracted))
    print("Percentage per neg root words in neg contexts extracted: ", [round(o/neg_contexts_extracted, 2) for o in neg_occ_extracted])
 