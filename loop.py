import requests
import time
import logging
import os
import gzip
import shutil
from check import *

import smtplib
from email.mime.text import MIMEText

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

import subprocess
import json

def send_simple_email(user_email, doc_id):
    """Send email with link to checked document"""
    try:
        # Email config - CHANGE THESE!
        sender = "lexiqo_app@mail.ru"
        password = ""
        smtp_server = "smtp.mail.ru"
        smtp_port = 465
        
        # Create link
        link = f"https://check.lexiqo.ru/results/{doc_id}"
        
        # Simple email
        msg = MIMEText(f"Your document is ready!\n\nView it here: {link}")
        msg['Subject'] = f"Document {doc_id} is ready"
        msg['From'] = sender
        msg['To'] = user_email
        
        # Send
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender, password)
        server.send_message(msg)
        server.quit()
        
        print(f"Email sent to {user_email}")
        result_warn = "Email is SENT"
        return result_warn
        
    except Exception as e:
        print(f"Email failed: {e}")
        result_warn = f"EMAIL IS NOT SENT: {e}"
        return result_warn

def docx_to_pdf(file):
    """Convert DOCX to PDF in same folder"""
    opts = json.dumps({
        "ExportNotes": {"type": "boolean", "value": "true"},
        "ExportBookmarks": {"type": "boolean", "value": "true"},
        "UseTaggedPDF": {"type": "boolean", "value": "true"}
    })
    out_dir = os.path.dirname(file) or '.'
    pdf_file = os.path.splitext(file)[0] + '.pdf'
    
    cmd = ['soffice', '--headless', '--convert-to', f'pdf:writer_pdf_Export:{opts}', 
           '--outdir', out_dir, file]
    
    try:
        subprocess.run(cmd, check=True, capture_output=True)
        print(f"✓ Created: {pdf_file}")
        return pdf_file
    except subprocess.CalledProcessError as e:
        print(f"✗ Error: {e.stderr.decode()}")
        return None

class DocProcessor:
    def __init__(self, api_url, auth_token, interval=60):
        self.api_url = api_url.rstrip('/')
        self.headers = {
            'Authorization': auth_token,
            'Content-Type': 'application/json'
        }
        self.interval = interval
        self.processed_count = 0

    def get_random_doc(self):
        try:
            response = requests.get(f"{self.api_url}/get_work/", headers=self.headers, timeout=30)
            if response.status_code == 200:
                data = response.json()
                if data.get("id") is not None:
                    logger.debug(f"Got document from API: {data['id']}")
                    return data
                else:
                    #logger.info("No documents available")
                    return None
            elif response.status_code == 403:
                logger.error("Access denied. Check your token and permissions.")
                return None
            else:
                logger.error(f"GET error {response.status_code}: {response.text}")
                return None
        except Exception as e:
            logger.error(f"Error getting document: {e}")
            return None

    def mark_as_processed(self, file_id, success=True, integer_string="0#0#0#0#0#0#0#0#0#0#0", plagiate=" "):
        try:
            payload = {"file_id": file_id, "processed": success, "integer_string": integer_string, "plagiate": plagiate}
            response = requests.post(f"{self.api_url}/save_work/", headers=self.headers, json=payload, timeout=30)
            if response.status_code == 200:
                logger.info(f"✓ Document {file_id} marked as processed")
                return True
            else:
                logger.error(f"✗ Failed to update document status: {response.status_code}")
                logger.error(f"Response: {response.text[:200]}")
                return False
        except Exception as e:
            logger.error(f"Error updating document status: {e}")
            return False

    def process_document(self, doc_data):
        try:
            doc_id = doc_data['id']
            style = doc_data.get('style', 'default')
            format_type = doc_data.get('format', 'standard')
            dictionary = doc_data.get('dictionary', 'stopwords')
            skip_pages_raw = doc_data.get('skip_pages', -1)

            user_email = doc_data.get('user', 'err')

            logger.info(f"Processing document {doc_id}")
            
            gz_path = os.path.join(os.getcwd(), 'djangoproject/static/docs', f'{doc_id}.docx.gz')
            docx_path = os.path.join(os.getcwd(), 'djangoproject/static/docs', f'{doc_id}.docx')
                
            try:
                with gzip.open(gz_path, 'rb') as gz_file, open(docx_path, 'wb') as out_file:
                    shutil.copyfileobj(gz_file, out_file)
                logger.debug(f"Extracted {gz_path} to {docx_path}")
            except Exception as e:
                logger.error(f"Error extracting {gz_path}: {e}")

            # Load dictionary
            dictionary_path = os.path.join(os.getcwd(), 'djangoproject/static/dictionary', f'{dictionary}.txt')
            if not os.path.exists(dictionary_path):
                logger.error(f"Dictionary file not found: {dictionary_path}")
                return False

            with open(dictionary_path, 'r', encoding='utf-8') as f:
                dict_words = f.read().splitlines()

            target_chars = [w for w in dict_words if len(w) == 1 and w.lower() != "я"]
            target_words = [w for w in dict_words if len(w) > 1 or w.lower() == "я"]

            logger.debug(f"Loaded dictionary with {len(target_chars)} single chars and {len(target_words)} words")

            xml_path = os.path.join(os.getcwd(), 'djangoproject/static/xml', f'{style}.xml')
            csl_path = format_type

            original_path = docx_path

            if not os.path.exists(gz_path):
                logger.error(f"File not found: {gz_path}")
                logger.info('Convert doc to docx')

                

                with gzip.open(os.path.join(os.getcwd(), 'djangoproject/static/docs', f'{doc_id}.doc.gz'), "rb") as f_in, open(os.path.join(os.getcwd(), 'djangoproject/static/docs', f'{doc_id}.doc'), "wb") as f_out:
                    shutil.copyfileobj(f_in, f_out)

                original_path = docx_to_pdf_simple(os.path.join(os.getcwd(), 'djangoproject/static/docs', f'{doc_id}.doc'), format_to='docx')

                logger.debug(f"Converted DOC to DOCX: {original_path}")
            
            # Convert to PDF with Notes

            skip_pages = 0

            try:
                skip_pages = int(skip_pages_raw)
            except Exception:
                logger.warning(f"Invalid skip_pages value '{skip_pages_raw}', defaulting to -1")
                skip_pages = 0

            end_of_n_page = -1

            if skip_pages > 0:
                logger.debug(f"skip_pages is set to {skip_pages}, starting PDF conversion and page mapping")

                get_pdf = docx_to_pdf_simple(original_path, format_to='pdf')
                if not get_pdf or get_pdf == 0 or not os.path.exists(str(get_pdf)):
                    logger.error(f"PDF conversion failed or file missing: {get_pdf}")
                    end_of_n_page = -1
                else:
                    logger.debug(f"PDF generated at {get_pdf}")

                    try:
                        pages_text = get_pdf_pages_text(get_pdf)

                        paragraphs_text = get_docx_paragraphs_text(original_path)

                        if pages_text and paragraphs_text:
                            page_map = map_paragraphs_to_pages_sequential(pages_text, paragraphs_text)
                            logger.debug(f"Created page map for {len(page_map)} pages")

                        else:
                            logger.warning("No pages or paragraphs found for mapping")
                            page_map = {}

                    except Exception as e:
                        logger.error(f"Exception during page mapping: {e}")
                        page_map = {}

                    
                    try:
                        end_of_n_page = max(page_map.get(skip_pages, [-1]))
                        logger.warning(f"Finally, end_of_n_page: {end_of_n_page}")
                    except:
                        logger.warning(f"Finally, end_of_n_page: {end_of_n_page}")
                        end_of_n_page = -1

            logger.warning(f"Finally, end_of_n_page: {end_of_n_page}")

            logger.info(f"Starting document analysis on {original_path}")
            output, plagiate, count_chars, count_words, count_sentences, count_bad_words, count_bad_chars, count_bibliography, count_bad_bibliography, count_not_doi, count_not_right_bibliography, count_styles_error, char_string, count_suggest_doi = get_document_properties(
                original_path,
                load_style_file=xml_path,
                target_chars=target_chars,
                target_words=target_words,
                do_spellcheck=True,
                do_styles=True,
                do_bibliography=True,
                bibliography_style=csl_path,
                skip_paras=end_of_n_page
            )

            #logger.info(plagiate)

            logger.warning(plagiate)

            list_of_integers = [
                count_words,
                count_chars,
                count_sentences,
                count_bad_words,
                count_bad_chars,
                count_bibliography,
                count_bad_bibliography,
                count_not_doi,
                count_suggest_doi,
                count_not_right_bibliography,
                count_styles_error,
            ]

            integer_string = "0#0#0#0#0#0#0#0#0#0#0"

            integer_string = "#".join([str(int(i)) for i in list_of_integers])
            logger.info(f'integer_string: {integer_string}')

            logger.info(f"Document processed: chars={count_chars}, words={count_words}")

            return True, integer_string, plagiate

        except Exception as e:
            logger.error(f"Unexpected error during document processing: {e}", exc_info=True)
            return False

    def run(self):
        logger.info("Starting document processor")
        logger.info(f"API URL: {self.api_url}")
        logger.info(f"Using token: {self.headers['Authorization'][:10]}...")
        logger.info(f"Polling interval: {self.interval} seconds when no documents")

        while True:
            try:
                doc_data = self.get_random_doc()
                if doc_data:
                    success, integer_string, plagiate = self.process_document(doc_data)
                    if self.mark_as_processed(doc_data['id'], success, integer_string, plagiate):
                        self.processed_count += 1
                        logger.info(f"Processed documents count: {self.processed_count}")
                    else:
                        logger.warning(f"Failed to mark document {doc_data['id']} as processed")
                else:
                    #logger.info(f"No documents found, sleeping for {self.interval}s")
                    time.sleep(self.interval)
            except KeyboardInterrupt:
                logger.info("Stopped by user")
                break
            except Exception as e:
                logger.error(f"Unexpected error in run loop: {e}", exc_info=True)
                time.sleep(self.interval)


if __name__ == "__main__":
    API_URL = "https://check.lexiqo.ru/api/v1"
    AUTH_TOKEN = ""
    
    processor = DocProcessor(API_URL, AUTH_TOKEN, interval=2)
    processor.run()