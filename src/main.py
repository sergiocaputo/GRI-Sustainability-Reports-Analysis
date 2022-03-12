import json
import os

from langdetect import detect
from GRIs import GRI
from sentiment_analysis import GRI_reports_to_excel_with_sentiment, compute_sentiment
from utils import get_text_from_pdf, preprocess_text, text_from_pdf, GRI_reports_to_excel, compute_metrics, merge_excels


def pdf_to_text(pdf_dir, texts_dir):
    """
    Transform PDF files into text documents using Tesseract OCR 
    :param pdf_dir: Path to the directory containing the PDF files to extract text documents
    :param texts_dir: Path to the directory where text documents will be saved
    """
    if not os.path.exists(texts_dir):
        os.makedirs(texts_dir)
     
    for p in os.listdir(pdf_dir):
        pdf_name = os.path.splitext(p)[0]
        path_text = os.path.join(texts_dir, pdf_name)
        os.makedirs(path_text)
        p = os.path.join(pdf_dir, p)
        print(p, path_text)
        text_from_pdf(p, path_text)


def extract_txt_from_pdf(pdf_dir, texts_dir):
    """
    Transform PDF files into text documents using PDFplumber
    :param pdf_dir: Path to the directory containing the PDF files to extract text documents
    :param texts_dir: Path to the directory where text documents will be saved
    """
    if not os.path.exists(texts_dir):
        os.makedirs(texts_dir)
     
    for p in os.listdir(pdf_dir):
        pdf_name = os.path.splitext(p)[0]
        path_text = os.path.join(texts_dir, pdf_name)
        os.makedirs(path_text)
        p = os.path.join(pdf_dir, p)
        print(p, path_text)
        get_text_from_pdf(p, path_text)


def create_GRI_json(report_texts, GRI_dir):
    """
    Create json file for each PDF report indicating the presence of GRI disclosures
    :param report_texts: Path to the directory containing the text documents previously extracted from PDF reports
    :param GRI_dir: Path to the directory where json files will be saved
    """
    keys = GRI.keys()
    if not os.path.exists(GRI_dir):
        os.makedirs(GRI_dir)

    for dir in os.listdir(report_texts):
        GRI_doc = GRI.copy()
        dir_path = os.path.join(report_texts, dir)

        for page in [f for f in os.listdir(dir_path) if f.endswith('.txt')]:
            for k in keys:
                if open(os.path.join(dir_path, page), 'r', errors="ignore").read().find(k) != -1:
                    GRI_doc[k] = page

        with open(os.path.join(GRI_dir,  "{}_GRI_report.json".format(dir)), 'w') as f:
            json.dump(GRI_doc, f, indent = 4) 
    

def extract_context(GRI_dir, texts):
    """
    Extract context for each GRI disclosure detected in text documents about PDF reports
    :param GRI_dir: Path to the directory containing the json files about GRI disclosures in reports
    :param texts: Path to the directory where text documents are stored
    """
    for doc in os.listdir(GRI_dir):
        doc_file = os.path.join(GRI_dir, doc)

        with open(doc_file) as f:
            json_GRI = json.load(f)

        print(doc_file)
        for k in GRI.keys():
            if json_GRI[k] != False:
                page = json_GRI[k]

                txt_file = os.path.join(os.path.join(texts, doc.split('_')[0]), page)

                with open(txt_file, 'r', errors="ignore") as txt:
                    text = txt.read().replace('\n', ' ')
                    context = text.partition(k)

                    context_before_gri = preprocess_text(' '.join(context[0].split()[-20:]))
                    context_after_gri = preprocess_text(' '.join(context[2].split()[:20]))

                json_GRI[k] = {'page': page, 'context_before': context_before_gri, 'context_after': context_after_gri}


                with open(doc_file, 'w') as f:
                    json.dump(json_GRI, f, indent = 4)


def extract_text():
    for texts_dir in ['texts_ocr', 'texts_extracted']:
        if texts_dir == 'texts_ocr':
            GRI_directory = 'GRI_reports_ocr'
            pdf_to_text(pdf_dir, texts_dir)
        else:
            GRI_directory = 'GRI_reports_extracted'
            extract_txt_from_pdf(pdf_dir, texts_dir)

        create_GRI_json(texts_dir, GRI_directory)
        extract_context(GRI_directory, texts_dir)


def create_excel_files():
    for excel_file in ['GRI_ocr.xlsx', 'GRI_extracted.xlsx']:
        if excel_file == 'GRI_ocr.xlsx':
            GRI_directory = 'GRI_reports_ocr'
        else:
            GRI_directory =  'GRI_reports_extracted'
    
        GRI_reports_to_excel(excel_file, GRI_directory)
        compute_metrics(excel_file, GRI_directory)
    
    merge_excels('GRI_ocr.xlsx', 'GRI_extracted.xlsx', 'GRI_final.xlsx')
    compute_metrics('GRI_final.xlsx', GRI_directory)


def create_excel_files_with_sentiment():
    for excel_file in ['GRI_ocr_sentiment.xlsx', 'GRI_extracted_sentiment.xlsx']:
        if excel_file == 'GRI_ocr_sentiment.xlsx':
            GRI_directory = 'GRI_reports_ocr'
        else:
            GRI_directory =  'GRI_reports_extracted'
    
        GRI_reports_to_excel_with_sentiment(excel_file, GRI_directory)
        compute_metrics(excel_file, GRI_directory)
    
    merge_excels('GRI_ocr_sentiment.xlsx', 'GRI_extracted_sentiment.xlsx', 'GRI_final_sentiment.xlsx')
    compute_metrics('GRI_final_sentiment.xlsx', GRI_directory)


def define_sentiment():
    for GRI_dir in ['GRI_reports_extracted', 'GRI_reports_ocr']:
        compute_sentiment(GRI_dir)


if __name__ == "__main__":

    pdf_dir = 'reports'

    #extract_text(pdf_dir)
    #create_excel_files()
    
    #define_sentiment()
    create_excel_files_with_sentiment()
  

