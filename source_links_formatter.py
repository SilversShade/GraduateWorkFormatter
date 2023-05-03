import re

from docx import Document


class SourceLinksFormatter:

    @staticmethod
    def find_reference_superscripts(doc: Document):
        for p in doc.paragraphs:
            for r in p.runs:
                if r.font.superscript is True and re.match(r"\[[0-9]+]", r.text):
                    print(r.text)

    @staticmethod
    def check_for_links_presence(doc: Document):
        SourceLinksFormatter.find_reference_superscripts(doc)
