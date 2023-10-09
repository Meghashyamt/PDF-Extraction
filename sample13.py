from PyPDF2 import PdfReader

def search_pdf_and_extract_nearby_text(pdf_file_path, keyword, num_lines=20):
    """Searches all pages of the given PDF document for the given keyword and extracts nearby text.

    Args:
        pdf_file_path: The path to the PDF file.
        keyword: The keyword to search for.
        num_lines: Number of lines to extract around the keyword (default is 20).

    Returns:
        A string containing the nearby text of the keyword found in any page.
    """
    # Open the PDF file.
    pdf_reader = PdfReader(pdf_file_path)

    nearby_text = ""
    # Iterate through all the pages of the PDF.
    for page in pdf_reader.pages:
        # Extract the text from the current page.
        text = page.extract_text()
        # Split the text into lines.
        lines = text.split("\n")
        # Find the index of the line containing the keyword.
        keyword_line_indices = [i for i in range(len(lines)) if keyword in lines[i]]

        # Extract nearby text for each occurrence of the keyword on the page.
        for keyword_line_index in keyword_line_indices:
            start_index = max(0, keyword_line_index - num_lines // 2)
            end_index = min(len(lines), keyword_line_index + num_lines // 2 + 1)
            nearby_text += " ".join(lines[start_index:end_index]) + "\n"

    return nearby_text

# Example usage:
pdf_file_path = r"C:\Users\M_Thiruveedula\Downloads\Projects_Code_Backup\Genus\RFP-214.pdf"
keyword = "Overview of the AMISP Scope of Work"

nearby_text = search_pdf_and_extract_nearby_text(pdf_file_path, keyword, num_lines=20)

print(nearby_text)

# from PyPDF2 import PdfReader

# def search_pdf_and_extract_nearby_text(pdf_file_path, keyword):
#     """Searches all pages of the given PDF document for the given keyword and extracts nearby text.

#     Args:
#         pdf_file_path: The path to the PDF file.
#         keyword: The keyword to search for.

#     Returns:
#         A string containing the nearby text of the keyword found in any page.
#     """
#     # Open the PDF file.
#     pdf_reader = PdfReader(pdf_file_path)

#     nearby_text = ""
#     # Iterate through all the pages of the PDF.
#     for page in pdf_reader.pages:
#         # Extract the text from the current page.
#         text = page.extract_text()
#         # Split the text into lines.
#         lines = text.split("\n")
#         # Find the index of the line containing the keyword.
#         keyword_line_indices = [i for i in range(len(lines)) if keyword in lines[i]]

#         # Extract nearby text for each occurrence of the keyword on the page.
#         for keyword_line_index in keyword_line_indices:
#             for line_index in range(keyword_line_index - 1, keyword_line_index + 2):
#                 if line_index >= 0 and line_index < len(lines):
#                     nearby_text += lines[line_index] + " "
    
#     return nearby_text

# # Example usage:
# pdf_file_path = r"C:\Users\M_Thiruveedula\Downloads\Projects_Code_Backup\Genus\RFP-214.pdf"
# keyword = "Overview of the AMISP Scope of Work"

# nearby_text = search_pdf_and_extract_nearby_text(pdf_file_path, keyword)

# print(nearby_text)
