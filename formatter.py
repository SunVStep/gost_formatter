from gost732.core.document_processor import DocumentProcessor
from pathlib import Path

def main():
    processor = DocumentProcessor()
    processor.process(
        input_path=Path("reports/test.docx"),
        template_path=Path("template/Main_Template.docx"),
        output_path=Path("output/result.docx")
    )

if __name__ == "__main__":
    main()