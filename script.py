from __future__ import print_function

import os
from dotenv import load_dotenv
load_dotenv()

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/documents']

# The ID of a sample document.
DOCUMENT_ID = #document id of the google doc
COMPANY_NAME = os.getenv("COMPANY_NAME")

ENV_VARS = {}
for k,v in os.environ.items():
    ENV_VARS.update({k:v})

def read_paragraph_element(element):
    """Returns the text in the given ParagraphElement.

        Args:
            element: a ParagraphElement from a Google Doc.
    """
    text_run = element.get('textRun')
    if not text_run:
        return ''
    return text_run.get('content')

def read_structural_elements(elements):
    """Recurses through a list of Structural Elements to read a document's text where text may be
        in nested elements.

        Args:
            elements: a list of Structural Elements.
    """
    text = ''
    for value in elements:
        if 'paragraph' in value:
            elements = value.get('paragraph').get('elements')
            for elem in elements:
                text += read_paragraph_element(elem)
        elif 'table' in value:
            # The text in table cells are in nested Structural Elements and tables may be
            # nested.
            table = value.get('table')
            for row in table.get('tableRows'):
                cells = row.get('tableCells')
                for cell in cells:
                    text += read_structural_elements(cell.get('content'))
        elif 'tableOfContents' in value:
            # The text in the TOC is also in a Structural Element.
            toc = value.get('tableOfContents')
            text += read_structural_elements(toc.get('content'))
    return text


def main():
    """Shows basic usage of the Docs API.
    Prints the title of a sample document.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('docs', 'v1', credentials=creds)
        
        title = 'Cover Letter'
        body = {
            'title': title
        }
        doc = service.documents().get(documentId=DOCUMENT_ID).execute()
        doc_content = doc.get('body').get('content')
        template = read_structural_elements(doc_content)
        template = template.replace('[','*[')
        template = template.replace(']',']*')
        words_list = template.split('*')
        
        for index, x in enumerate(words_list):
            keywords=x[x.find("[")+1:x.find("]")]
            if keywords in ENV_VARS:
                words_list[index] = ENV_VARS[keywords]
                
        filled_template = "".join(words_list)
        
        with open(f"{COMPANY_NAME}.txt", "w") as file:
            # Writing data to a file
            file.write(filled_template)
        
    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()