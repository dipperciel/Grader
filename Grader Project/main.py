# --------------------------------------------------------------------------------------------------- #
# Grading module
# --------------------------------------------------------------------------------------------------- #

# Import packages ----------------------------------------------------------------------------------- #
from docx import Document # install python-docx
import re
# import pyapa

# INPUT variables------------------------------------------------------------------------------------ #

doc = Document("C:/Users/donal/OneDrive - York University/New/Al/E-Grader/test_apa.docx") #Document to be graded
end_of_paragraph = r"\.\s*$" #Regex pattern to find the end of a paragraph (a hard return)

# Functions------------------------------------------------------------------------------------------ #

def create_three_sections(doc, end_of_paragraph): #  Splits the text into three sections: title page, body and references
    # start with identifying the three parts of the document: title page, body and references
    text = ""  # Variable to hold the full text of the document
    found_beginning = False # Variable to check if the beginning of the post-title text portion of the document has been found
    for paragraph in doc.paragraphs:
        if re.search(end_of_paragraph, paragraph.text) and found_beginning == False: # If there is a real end of a paragraph
            text += "*Beginning*" + paragraph.text + "\n"
            found_beginning = True  # Set the flag to true to indicate that the beginning of the text portion of the document has been found and marked
        else:
            text += paragraph.text + "\n"  # Add the text of each paragraph to the variable

    full_text = text  # Create full text of document
    index1 = full_text.find("*Beginning*")
    title_page = full_text[:index1 - 1]  # Create title page
    text_minus_title = full_text[index1 + 11:]  # Create text minus title page

    index2 = full_text.find("References\n")  # Find the index of the reference section, which has the title "References"
    if index2 == -1:  # If there is no reference section
        index2 = full_text.find(
            "Bibliography\n")  # Find the index of the reference section, which has the title "Bibliography"
        if index2 == -1:  # If there is no Biblioraphy section
            references = "null"
            body = full_text[index1 + 11:index2 - 1]  # Create body section
        else:
            references = full_text[index2 + 12:]  # Create reference section starting with "Bibliography"
            body = full_text[index1+11:index2 - 1]  # Create body section
    else:
        references = full_text[index2 + 10:]  # Create references section starting with "References"
        body = full_text[index1+11:index2 - 1]  # Create body section

    full_text = full_text.replace("*Beginning*", "")

    return title_page, text_minus_title, references, body, full_text

def wordcount(text): #  Returns the number of words in a string
    word_list = text.split()  # Split the text into a list of words
    word_count = len(word_list)  # Count the number of words in the list
    return word_count


# Main ---------------------------------------------------------------------------------------------- #
# --------------------------------------------------------------------------------------------------- #

title_page, text_minus_title, references, body, full_text = create_three_sections(doc, end_of_paragraph) # Creation of title page, text minus title page and reference from the document

# First have to test if there's a bibliography. If not, it should be added to the report !!!!! And bypass References check

# Word counts --------------------------------------------------------------------------------------- #
num_words_text_minus_title = wordcount(text_minus_title)
num_body = wordcount(body)
num_words_full_text = wordcount(full_text) - 1 #Count the number of words in the full text of the document. -1 is to remove "*Beginning*" [SHOULD I DELETE THE WORD?*****]

# Check APA style ----------------------------------------------------------------------------------- #
error_references = []
if references != "null":
    reference_list = references.splitlines()  # Get the individual references in a list
    reference_list = [item for item in reference_list if item]
    num_references = len(reference_list)

    index1 = reference_list[0].index(").") # testing

    ref_author_year_part = reference_list[0][:index1+2]
    print(ref_author_year_part)
else:
    error_references.append("There is no references section")

# find if thers's leading space in ref_author_year_part. If not: Message error: Missing space between year and title. !!!

# In-text citations
# pattern = r"\(([^)]*), (\d{4})(?:, p\. (\d{1,4}))?\)" # e.g. (Ipperciel, 2018, p. 12). THIS IS TOO PRECISE. WON'T PICK UP ON MISTAKES
pattern = r"\([^)]*\d{4,}[^)]*\)"
citation_list = re.findall(pattern, body)
print("citation_list: ", citation_list)

# First check if there are three parts in the matches!!! If not: missing comma after author

error_APA = []
i = 0
for citation in citation_list:
    if not re.search(", \d{4}", citation):  # first check if there's a comma and space before the year
        error_APA.append("Missing comma before year in reference " + str(i+1) + ": " + citation_list[i])
    if re.search("\d{4}, \d{1,2,3,4}", citation):  # check if there's a space and comma after the year, and only a number instead of p.
        error_APA.append("Missing 'p. ' before page number in reference " + str(i+1) + ": " + citation_list[i])
    if "et al" in citation_list[i]:  # Check "et al" mistakes
        if "et al." not in citation_list[i]:
            error_APA.append("Missing period after 'et al' in reference " + str(i+1) + ": " + citation_list[i])
        if ", et al" in citation_list[i] or ",et al." in citation_list[0]:
            error_APA.append("No comma should precede 'et al.' (the comma always follows 'et al.'). See reference " + str(i+1) + ": " + citation_list[i])
    if " and " in citation_list[i]:  # Check if "and" instead of & is in the citation
        error_APA.append("Ampersand (&) should be used instead of 'and' in reference " + str(i + 1) + ": " + citation_list[i])
    # Next: check if authors are juxtaposed with comma instead of & !!!!

    i = i + 1
print(error_APA)

# Do citations match with the references? !!!!!!!

# Author check
references = references.lstrip() #Remove  whitespace from the references section
test = "Ipperciel, D., Elatia, S. (2018a)."





#### Define regex patterns for APA ####
APA_author = r"^\b([A-Z][a-z]*),\s([A-Z])\." # e.g. Ipperciel, D.
APA_author_next = r"\,\s([A-Z][a-z]*),\s([A-Z])\.\s" # e.g. & ElAtia, S.
APA_author_all = f"{APA_author}(?:{APA_author_next})*" # e.g. Ipperciel, D. & ElAtia, S.
APA_year = r"\([0-9]{4}[a-z]?\)\."

### Checking accuracy of author format
#if re.search(APA_author, test): #If the test string matches the pattern
#    print("Author format is correct")
if re.search(APA_author_all,test):
    print("Author format is also correct")
else: #If the test string does not match the pattern
    print("Author format is incorrect")
