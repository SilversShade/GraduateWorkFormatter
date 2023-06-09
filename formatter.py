from docx import Document
from os import path

from main_req_formatter import MainRequirementsFormatter
from source_links_formatter import SourceLinksFormatter
import argparse


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('path_to_docx', type=str, help='Path to docx file')
    args = parser.parse_args()

    if not path.exists(args.path_to_docx) or not path.isfile(args.path_to_docx):
        print('Введен неверный путь до файла')
        exit(1)

    try:
        doc = Document(args.path_to_docx)
    except ValueError:
        print('Документ должен быть типа docx')
        exit(1)

    MainRequirementsFormatter.format_document(doc)
    MainRequirementsFormatter.change_title_page_year(doc, '2023')
    SourceLinksFormatter.check_for_links_presence(doc)

    doc.save("edited.docx")


if __name__ == '__main__':
    main()
