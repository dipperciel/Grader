# --------------------------------------------------------------------------------------------------- #
# Grading module
# --------------------------------------------------------------------------------------------------- #

# Import packages ----------------------------------------------------------------------------------- #
from docx import Document  # install python-docx
import re

# import neuspell
# from neuspell import available_checkers, BertChecker

# INPUT variables------------------------------------------------------------------------------------ #

doc = Document(
    "C:/Users/donal/OneDrive - York University/New/Al/E-Grader/test_references.docx")  # Document to be graded
end_of_paragraph = r"\.\s*$"  # Regex pattern to find the end of a paragraph (a hard return)
required_wordcount = 1000
required_references = 5


# Functions------------------------------------------------------------------------------------------ #

def create_three_sections(doc,
                          end_of_paragraph):  # Splits the text into three sections: title page, body and references
    # start with identifying the three parts of the document: title page, body and references
    text = ""  # Variable to hold the full text of the document
    found_beginning = False  # Variable to check if the beginning of the post-title text portion of the document has been found
    for paragraph in doc.paragraphs:
        if re.search(end_of_paragraph,
                     paragraph.text) and found_beginning == False:  # If there is a period and hard return, that's the first paragraph of the body
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
            body = full_text[index1 + 11:index2 - 1]  # Create body section
    else:
        references = full_text[index2 + 10:]  # Create references section starting with "References"
        body = full_text[index1 + 11:index2 - 1]  # Create body section

    full_text = full_text.replace("*Beginning*", "")

    return title_page, text_minus_title, references, body, full_text


def wordcount(text):  # Returns the number of words in a string
    word_list = text.split()  # Split the text into a list of words
    word_count = len(word_list)  # Count the number of words in the list
    return word_count


def check_intext_citations(body):
    # pattern = r"\(([^)]*), (\d{4})(?:, p\. (\d{1,4}))?\)" # e.g. (Ipperciel, 2018, p. 12). THIS IS TOO PRECISE. WON'T PICK UP ON MISTAKES
    pattern = r"\([^)]*[1,2]\d{3}[a-zA-Z]?[^)]*\)"
    citation_list = re.findall(pattern, body)  # finds all in-text citations in the body
    # print("citation_list: ", citation_list)

    error_APA = []
    i = 0
    # all rules for in-text citations in if clauses
    for citation in citation_list:
        if not re.search(", [1,2]\d{3}", citation):  # check if there's a comma and space before the year
            error_APA.append("Missing comma before year in reference " + str(i + 1) + ": " + citation_list[i])
        if re.search("[1,2]\d{3}[a-z]?, \d{1,4}", citation) or re.search("[1,2]\d{3}[a-z]?,\d{1,4}",
                                                                         citation):  # check if there's a comma after the year, and only a number instead of p.
            error_APA.append("Missing 'p. ' before page number in reference " + str(i + 1) + ": " + citation_list[i])
        if not re.search("[1,2]\d{3}[a-z]?[,)]", citation):  # check if there's a missing comma after the year
            error_APA.append("Missing comma after the year in reference " + str(i + 1) + ": " + citation_list[i])
        if "et al" in citation_list[i]:  # Check "et al" mistakes
            if "et al." not in citation_list[i]:
                error_APA.append("Missing period after 'et al' in reference " + str(i + 1) + ": " + citation_list[i])
            if ", et al" in citation_list[i] or ",et al." in citation_list[0]:
                error_APA.append(
                    "No comma should precede 'et al.' (the comma always follows 'et al.'). See reference " + str(
                        i + 1) + ": " + citation_list[i])
        if " and " in citation_list[i]:  # Check if "and" instead of & is in the citation
            error_APA.append(
                "Ampersand (&) should be used instead of 'and' in reference " + str(i + 1) + ": " + citation_list[i])
        # Next: check if authors are juxtaposed with comma instead of & !!!!

        i = i + 1

    # Check for page number !!!

    return error_APA, citation_list


def check_author_year(reference_list, ref_author_year_part, index1, i, error_references):

    # Possible errors
    if reference_list[i][index1 + 1] != ".":
        error_references.append(
            "Missing period after the bracketed year in Reference " + str(i + 1) + ": " + reference_list[i][
                                                                                          :index1 + 1])
    if " and " in ref_author_year_part[i]:  # Check if "and" instead of & is in the citation
        error_references.append(
            "Ampersand (&) should be used instead of 'and' in reference " + str(i + 1) + ": " +
            ref_author_year_part[i])

    # Preparing the author section for analysis below
    temp_item = ref_author_year_part[i].replace("&", ",").replace("and", ",").split(",")
    temp_item[-1] = temp_item[-1][:-6]
    temp_item = [x for x in temp_item if
                 x != "" and x != " "]  # Removes the items with the values "" and " " from the list

    # Analysis based on the number of author parts (items separated by a comma)
    if len(temp_item) == 1:  # if there's no comma in the single-author reference
        error_references.append(
            "Missing comma between author name and first name in Reference " + str(i + 1) + ": " +
            ref_author_year_part[i])
        if temp_item[0][-1] != " ":
            error_references.append(
                "Missing space between author and bracketed year in Reference " + str(i + 1) + ": " +
                ref_author_year_part[i])
        if temp_item[0].rstrip()[-1] != ".":
            error_references.append(
                "In Reference " + str(
                    i + 1) + ", the first name should appear after the author's name as a single capital letter followed by a period.")

    if len(temp_item) == 2:  # single author reference
        if not temp_item[1].endswith(" "):
            error_references.append(
                "Missing space between author and bracketed year in Reference " + str(i + 1) + ": " +
                ref_author_year_part[i])
        if not temp_item[1].startswith(" "):
            error_references.append(
                "Missing space between author name and first name in Reference " + str(i + 1) + ": " +
                ref_author_year_part[i])
        if not temp_item[1].rstrip().endswith("."):
            error_references.append(
                "In Reference " + str(
                    i + 1) + ", the first name should be a single capital letter followed by a period.")

    if len(temp_item) == 3:
        error_references.append(
            "Missing comma in author name in Reference " + str(i + 1) + ": " +
            ref_author_year_part[i])

    # If temp_item has on odd number of element, it means there's a comma missing. It has to be added in the right place
    if len(temp_item) % 2 != 0 and len(temp_item) > 3:
        error_references.append(
            "Missing comma in author name in Reference " + str(i + 1) + ": " + ref_author_year_part[i])
        for item in temp_item:
            if len(item.split()) == 2 and not item.split()[0][-1] == ".":
                last_name, first_name = item.split()
                first_name = " " + first_name
                temp_item = [last_name] + [first_name] + temp_item[1:]

    if len(temp_item) >= 4:  # Two or more authorsThree-author reference
        if "&" not in ref_author_year_part[i]:
            error_references.append(
                "In Reference " + str(i + 1)+ ", the last author should be separated from the rest by a comma and an ampersand (&).")
        counter = 0
        for item in temp_item[
                    1:]:  # the loop starts on the second items, as the first (ie author) will never start with a space
            counter += 1
            if not item.startswith(" "):
                error_references.append("Missing space in Reference " + str(i + 1) + ", at or after author " + str(
                    int((counter + 1) / 2)) + ": " + ref_author_year_part[i])
        if temp_item[len(temp_item) - 3].endswith(" "):
            error_references.append(
                "Missing comma before the ampersand or 'and' in Reference " + str(i + 1) + ": " + ref_author_year_part[
                    i])
        for j in range(1, len(temp_item), 2):
            if not re.search("[A-Z]\.", temp_item[j]):
                error_references.append("In Reference " + str(i + 1) + ", author " + str(
                    int((j + 1) / 2)) + ", the first name should be a single capital letter followed by a period.")

    return error_references


def check_references(references, required_references):
    error_references = []

    # Determine if the number of references meets the requirement
    reference_list = references.splitlines()  # Get the individual references in a list
    reference_list = [item for item in reference_list if item]
    num_references = len(reference_list)
    if num_references < required_references:
        error_references.append(
            "There are only " + str(num_references) + " references when the requirement is for " + str(
                required_references) + ".")

    # Check if the references are in alphabetical order
    if reference_list != sorted(reference_list):
        error_references.append("References in the bibliography are not in alphabetical order.")

    # Check if the references has a doi
    for i in range(len(reference_list)):
        if "https://doi" in reference_list[i]:
            error_references.append("Reference " + str(
                i + 1) + " has a doi number. While this is optional in APA, it should not be included in papers in this class.")

    # Check for missing period at the end of the reference
    for i in range(len(reference_list)):
        if not reference_list[i].endswith("."):
            error_references.append("Missing period at the end of reference " + str(i + 1) + ".")

    # Check inside the references-----------------------------------------------------------------#

    # Grab the names and authors from each reference
    ref_author_year_part = []
    for i in range(len(reference_list)):
        if re.search("\([1-2]\d{3}[a-z]?", reference_list[i]) or "(n.d.)" in reference_list[i]:  # If the reference is in APA, with a year in brackets

            # CASE 1 in APA: classic year in brackets, i.e. Ipperciel, D. (2023) or ElAtia, S. (2022a)
            if re.search("\([1-2]\d{3}[a-z]?\)", reference_list[i]):
                index1 = reference_list[i].index(")")
                ref_author_year_part.append(reference_list[i][:index1 + 1])

                check_author_year(reference_list, ref_author_year_part, index1, i, error_references)


            # CASE 2 in APA: Web reference, e.g. Ipperciel, D. (2023, October 31).
            elif re.search("\([1-2]\d{3}[a-z]?,\s?\w+ \d{1,2}\)", reference_list[i]):
                index1 = reference_list[i].index(")")
                temp_reference_list = reference_list[i][:index1 + 1] # Need to remove the month/day from the bracket for standardized treatment in check_author_year
                bracket_index = temp_reference_list.index("(")
                comma_index = temp_reference_list.index(",", bracket_index + 1)
                temp_reference_list = temp_reference_list[:comma_index] + ")"
                ref_author_year_part.append(temp_reference_list)

                check_author_year(reference_list, ref_author_year_part, index1, i, error_references)

                months = ["January", "February", "March", "April", "May", "June", "July", "August", "September",
                          "October", "November", "December"]
                if not any(month in reference_list[i][:index1 + 1] for month in months):
                    error_references.append("In Reference " + str(i + 1) + ", the month has to be written in full")

            # CASE 3: No year (n.d.)
            elif "(n.d.)" in reference_list[i]:
                index1 = reference_list[i].index(")")
                ref_author_year_part.append(reference_list[i][:index1 + 1])

                check_author_year(reference_list, ref_author_year_part, index1, i, error_references)

        else:  # If the reference is not in APA style, i.e. no year in brackets
            error_references.append("Reference " + str(i + 1) + " is not in APA format. References in APA always start with the author's name, first letter of the first name and the year in brackets.")
            temp_ref = reference_list[i].split()[0] + " (" + re.search("[1,2]\d{3}", reference_list[
                i]).group() + ")"  # grabs the first author's name and a year anywhere in the reference
            ref_author_year_part.append(temp_ref)

    #### Define regex patterns for APA ####
    # APA_author = r"^\b([A-Z][a-z]*),\s([A-Z])\."  # e.g. Ipperciel, D.
    # APA_author_next = r"\,\s([A-Z][a-z]*),\s([A-Z])\.\s"  # e.g. & ElAtia, S.
    # APA_author_all = f"{APA_author}(?:{APA_author_next})*"  # e.g. Ipperciel, D. & ElAtia, S.
    # APA_year = r"\([0-9]{4}[a-z]?\)\."

    # !!!! If there's an http (not doi), it should be only for web sites. Need to be able to identify a web page... so as to exclude http from other references !!!

    return error_references


def concordance_btw_citations_and_references(references,
                                             citation_list):  # Checks if the citations in the body match the references and visa versa
    error_concordance = []

    # Get the individual references in a list
    reference_list = references.splitlines()
    reference_list = [item for item in reference_list if item]
    # num_references = len(reference_list)

    # Grab the names and authors from each reference
    ref_author_year_part = []
    for i in range(len(reference_list)):
        try:
            index1 = reference_list[i].index(")")
            if not re.search("\([1,2]\d{3}[a-z]?\)", reference_list[i][
                                                     :index1 + 1]):  # in case the reference is not in APA format, a bracket may appear further downthe reference...
                temp_ref = reference_list[i].split()[0] + " (" + re.search("[1,2]\d{3}", reference_list[
                    i]).group() + ")"  # grabs the first author's name and a year anywhere in the reference
                ref_author_year_part.append(temp_ref)
            else:
                ref_author_year_part.append(reference_list[i][:index1 + 1])
        except:  # in case the references are not in APA format
            error_references.append(
                "In APA, references always start with the author's name, first letter of the first name and the year in brackets. See Reference " + str(
                    i + 1) + ".")
            temp_ref = reference_list[i].split()[0] + " " + re.search("[1,2]\d{3}", reference_list[
                i]).group()  # grabs the first author's name and a year anywhere in the reference
            ref_author_year_part.append(temp_ref)

    # The ref_author_year_part list must be cleaned up to look like a proper in-text citation
    cleaned_ref_author_year = []
    for i in range(len(ref_author_year_part)):
        temp_item = ref_author_year_part[i].split()
        if len(temp_item) == 2:  # Ipperciel (2010) -- 2 parts
            cleaned_ref_author_year.append(temp_item[0] + " " + temp_item[1][1:-1])
        elif len(temp_item) == 3:  # if the reference only has one author, e.g. Ipperciel, D. (2010) -- 3 parts
            cleaned_ref_author_year.append(temp_item[0] + " " + temp_item[2][1:-1])
        elif len(temp_item) == 6:  # if the reference has two authors, e.g. Ipperciel, D. & Elatia, S. (2010) -- 6 parts
            cleaned_ref_author_year.append(
                temp_item[0][:-1] + " " + temp_item[2] + " " + temp_item[3] + " " + temp_item[5][1:-1])
        elif len(
                temp_item) > 6:  # if the reference has more than two authors, e.g. Ipperciel, D., Elatia, S. & Johnson, M. (2010)
            cleaned_ref_author_year.append(temp_item[0][:-1] + " et al., " + temp_item[-1][1:-1])

    # Compare the citation_list with cleaned references
    for i in range(len(citation_list)):
        for j in range(len(cleaned_ref_author_year)):
            citation_year = re.search("[1,2]\d{3}", citation_list[i]).group()
            if (citation_list[i].split(",")[0].strip()[1:] not in cleaned_ref_author_year[
                j]  # first condition: first word in citation_list is in reference
                    or citation_year not in cleaned_ref_author_year[
                        j]):  # second condition: the year in citation is in reference
                in_the_references = 0
            else:
                in_the_references = 1
                break
        if in_the_references == 0:
            error_concordance.append("Citation " + str(i + 1) + " " + citation_list[i] + " is not in the references.")

    # Compare the cleaned references with the citation_list (there must be at least one citation for each reference)
    for i in range(len(cleaned_ref_author_year)):
        in_the_citations = 0
        for j in range(len(citation_list)):
            try:
                ref_year = re.search("[1,2]\d{3}", cleaned_ref_author_year[i]).group()
            except:
                ref_year = ""
            if cleaned_ref_author_year[i].split(",")[0].strip() in citation_list[j] and ref_year in citation_list[j]:
                in_the_citations = 1
                break
        if in_the_citations == 0:
            error_concordance.append("Reference " + str(i + 1) + " " + cleaned_ref_author_year[
                i] + " is not used as a citation in the body.")

    return error_concordance


def generate_final_report(required_wordcount, num_words_text_minus_title, error_APA,
                          error_references):  # !! pas oublier d'ajouter error_concordance!!!!
    final_report = ""
    if num_words_text_minus_title < (required_wordcount * 0.9):
        final_report = final_report + "Your text has fewer words than the required word count (i.e. " + str(
            required_wordcount) + " words minus 10% or " + str(int(required_wordcount * 0.9)) + " words).\n"
    if num_words_text_minus_title > (required_wordcount * 1.1):
        final_report = final_report + "Your text has more words than the required word count (i.e. " + str(
            required_wordcount) + " words plus 10% or " + str(int(required_wordcount * 1.1)) + " words).\n"
    if error_APA:
        final_report = final_report + "APA in-text citation error(s): \n"
        for i in range(len(error_APA)):
            final_report = final_report + error_APA[i] + "\n"

    if error_references:
        final_report = final_report + "Bibliography error(s): \n"
        for i in range(len(error_references)):
            final_report = final_report + error_references[i] + "\n"

    # if error_concordance:
    #     for i in range(len(error_concordance)):
    #         final_report = final_report + error_concordance[i] + "\n"

    return final_report


# Main ---------------------------------------------------------------------------------------------- #
# --------------------------------------------------------------------------------------------------- #

# Creation of title page, text minus title page and reference from the document
title_page, text_minus_title, references, body, full_text = create_three_sections(doc, end_of_paragraph)

# Word counts --------------------------------------------------------------------------------------- #
num_words_text_minus_title = wordcount(text_minus_title)
num_body = wordcount(body)
num_words_full_text = wordcount(
    full_text) - 1  # Count the number of words in the full text of the document. -1 is to remove "*Beginning*" [SHOULD I DELETE THE WORD?*****]

# Error report for in-text citations ----------------------------------------------------------------- #
error_APA, citation_list = check_intext_citations(body)
# print(error_APA)

# Are the in-text citations in references and vice versa? -------------------------------------------- #
# error_concordance = concordance_btw_citations_and_references(references, citation_list)

# Error report for references ------------------------------------------------------------------------ #
error_references = ""
if references != "null":  # Do this only if there is a references section
    error_references = check_references(references, required_references)
else:
    error_references = error_references + "There is no references section"
# print("Error report for references: ", error_references)
# print("Number of references: ", num_references)

# Generate final report ------------------------------------------------------------------------------ #

final_report = generate_final_report(required_wordcount, num_words_text_minus_title, error_APA,
                                     error_references)  # !! pas oublier d'ajouter error_concordance!!!!
print("Final report: ", final_report)
