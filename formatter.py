from docx import Document

from main_req_formatter import MainRequirementsFormatter
from source_links_formatter import SourceLinksFormatter
import argparse


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('path_to_docx', type=str, help='Path to docx file')
    args = parser.parse_args()

    doc = Document(args.path_to_docx)

    MainRequirementsFormatter.format_document(doc)
    MainRequirementsFormatter.change_title_page_year(doc, '2023')
    SourceLinksFormatter.check_for_links_presence(doc)

    doc.save("edited.docx")


if __name__ == '__main__':
    main()
