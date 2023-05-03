from docx import Document

from main_req_formatter import MainRequirementsFormatter
from source_links_formatter import SourceLinksFormatter


def main():
    doc = Document("test.docx")
    MainRequirementsFormatter.format_document(doc)
    SourceLinksFormatter.check_for_links_presence(doc)
    doc.save("edited.docx")


if __name__ == '__main__':
    main()
