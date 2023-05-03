from docx import Document


class SourceLinksFormatter:

    @staticmethod
    def check_for_links_presence(doc: Document):
        pass  # regex to match a number in square brackets: /\[[0-9]+\]/
