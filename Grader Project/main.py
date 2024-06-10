# --------------------------------------------------------------------------------------------------- #
# Grading module
# --------------------------------------------------------------------------------------------------- #
import json

# Import packages ----------------------------------------------------------------------------------- #
from docx import Document  # install python-docx
import re
import mammoth  # to convert docx to html
import urllib.request
import json
# import neuspell
# from neuspell import available_checkers, BertChecker

# INPUT variables------------------------------------------------------------------------------------ #
file_source = "C:/Users/donal/OneDrive - York University/New/Al/E-Grader/test_references.docx"
doc = Document(file_source)  # Document to be graded
end_of_paragraph = r"\.\s*$"  # Regex pattern to find the end of a paragraph (a hard return)
required_wordcount = 1000
required_references = 5

# Initial set-up

with open(file_source, "rb") as docx_file:
    result = mammoth.convert_to_html(docx_file)  # this converts my doc to html
    html_text = result.value

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


def create_html_references(html_text): # Creates a references section in html to check for italics
    if html_text.find(">Bibliography<") != -1:
        biblio_index = html_text.find(">Bibliography<")
        html_references = "<p" + html_text[biblio_index:]

    elif html_text.find(">References<") != -1:
        biblio_index = html_text.find(">References<")
        html_references = "<p" + html_text[biblio_index:]
    else:
        html_references = "null"

    return html_references


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


def check_doi(reference_list_i, error_references):

    if "https://doi.org/" in reference_list_i:
        # extract the doi link from the reference
        doi_link = re.search("https:\/\/doi\.org\/10\.\d+\S*[^.]", reference_list_i).group()[16:]
        crossref_url = "https://api.crossref.org/works/" + doi_link # this is the api.crossref.org url
        # read the contents of the URL using the urllib.request library (see import section)
        response = urllib.request.urlopen(crossref_url)
        content = response.read().decode("utf-8")
        # convert content to a true JSON file (it's already in a JSON format)
        data = json.loads(content)

        # Access specific elements
        reference_type = data["message"]["type"]
        # print("reference_type: " + reference_type)
        if reference_type == "journal-article":
            # print(" entered!!!!!")
            article_title = data["message"]["title"][0]
            journal_title = data["message"]["container-title"][0]
            journal_volume = data["message"]["volume"]
            journal_issue = data["message"]["journal-issue"]["issue"]
            article_pages = data["message"]["container-title"][0]
            # print("reference_type: ", reference_type)
            # print("article_title: ", article_title)
            # print("journal_title: ", journal_title)
            # print("journal_volume: ", journal_volume)
            # print("journal_issue: ", journal_issue)
            # print("article_pages: ", article_pages)
        else:
            pass

        # possible errors
        # list them here

    return error_references


def add_doi(reference_list, error_references):  # Add doi to references that don't have one
    authors = []
    titles = []
    for i in range(len(reference_list)):

        # grab first author name in reference_list and put it in author[]
        name_index = reference_list[i].index(",")
        temp_author = reference_list[i][:name_index]
        authors.append(temp_author)

        # find the title of the article or book or whatever
        bracket_index = reference_list[i].index(")") # !!!! Author can have first names in brackes !!!
        period_index = reference_list[i].index(".", bracket_index + 2)
        temp_title = reference_list[i][bracket_index+2:period_index].strip()
        temp_title_mod = temp_title.replace(" ", "+")
        titles.append(temp_title_mod)

        # search crossref using author and title
        crossref_url = "https://api.crossref.org/works?query.bibliographic=" + titles[i] + ".author=" + authors[i]
        encoded_url = urllib.parse.quote(crossref_url, safe=':/?=&') # so it can handle non-ASCII characters
        response = urllib.request.urlopen(encoded_url)
        content = response.read().decode("utf-8")
        data = json.loads(content)

        # Grab DOI from crossref
        doi = data["message"]["items"][0]["DOI"]
        clean_doi = doi.replace(".supp", "")
        doi = clean_doi

        # Check if doi is valid
        crossref_url = "https://api.crossref.org/works/" + doi  # this is the api.crossref.org url
        # read the contents of the URL using the urllib.request library (see import section)
        response = urllib.request.urlopen(crossref_url)
        content = response.read().decode("utf-8")
        # convert content to a true JSON file (it's already in a JSON format)
        data = json.loads(content)

        # Add the doi to the reference if the JSON data is in the reference list
        if "https://doi.org/" not in reference_list[i]:
            temp_reference_list = reference_list[i] + " https://doi.org/" + doi
        else:
            temp_reference_list = reference_list[i]
        print("temp_reference_list: ", temp_reference_list)
        # Compare specific elements
        reference_type = data["message"]["type"]
        print("reference_type: " + reference_type)
        if (data["message"]["title"][0] in temp_reference_list and
                data["message"]["author"][0]["family"] in temp_reference_list):
            reference_list[i] = temp_reference_list

    return reference_list


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

    # Check the author-year part of the references -----------------------------------------------#

    # Grab the names and authors from each reference
    ref_author_year_part = []
    for i in range(len(reference_list)):
        if re.search("\([1-2]\d{3}[a-z]?", reference_list[i]) or "(n.d.)" in reference_list[i]:  # If the reference is in APA, with a year in brackets

            # CASE 1 in APA: classic year in brackets, i.e. Ipperciel, D. (2023) or ElAtia, S. (2022a)
            if re.search("\([1-2]\d{3}[a-z]?\)", reference_list[i]):
                index1 = reference_list[i].index(")") # !!!! Author can have first names in brackes !!!
                ref_author_year_part.append(reference_list[i][:index1 + 1])

                check_author_year(reference_list, ref_author_year_part, index1, i, error_references)


            # CASE 2 in APA: Web reference, e.g. Ipperciel, D. (2023, October 31).
            elif re.search("\([1-2]\d{3}[a-z]?,\s?\w+ \d{1,2}\)", reference_list[i]):
                index1 = reference_list[i].index(")") # !!!! Author can have first names in brackes !!!
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
                index1 = reference_list[i].index(")") # !!!! Author can have first names in brackes !!!
                ref_author_year_part.append(reference_list[i][:index1 + 1])

                check_author_year(reference_list, ref_author_year_part, index1, i, error_references)

        else:  # If the reference is not in APA style, i.e. no year in brackets
            error_references.append("Reference " + str(i + 1) + " is not in APA format. References in APA always start with the author's name, first letter of the first name and the year in brackets.")
            temp_ref = reference_list[i].split()[0] + " (" + re.search("[1,2]\d{3}", reference_list[
                i]).group() + ")"  # grabs the first author's name and a year anywhere in the reference
            ref_author_year_part.append(temp_ref)

    # Check the second part of the references -----------------------------------------------------#

    html_references = create_html_references(html_text)  # This gives me the bibliography in html format
    html_reference_list = html_references.split("</p>")  # The bibliography is split into individual references
    html_reference_list = [item[3:-4] for item in html_reference_list if item != ''] # deletes empty items

    # Isolates the text in italics
    italicized = []
    for individual_reference in html_reference_list:
        if re.search("<em>.*?</em>", individual_reference):
            temp_italics = re.search("<em>.*?</em>", individual_reference).group()[4:-5]  # Grabs the content between the <em> tags
            temp_italics = re.sub(r",[\d\s()]*", "", temp_italics).strip()  # eliminates digits, spaces and parenthesis
            temp_italics = re.sub(r"\.[\d\s()]*","", temp_italics).strip() # in case the journal title is erroneously followed by a period
            italicized.append(temp_italics)
        elif "Bibliography<" in individual_reference or "References<" in individual_reference:
            pass
        else:
            italicized.append("")

    #print(italicized)
    #print(reference_list)

    # Add doi to references that don't have one
    reference_list = add_doi(reference_list, error_references)
    print("modified reference list: ", reference_list)

    for i in range(len(reference_list)):
        if italicized[i] != "":
            begin_index = reference_list[i].index(italicized[i])
            end_index = begin_index + len(italicized[i])
            before_italics = reference_list[i][:begin_index]
            after_italics = reference_list[i][end_index:]

            # Identify types of references ########
            error_references = check_doi(reference_list[i], error_references)

            # Journal articles
            if not re.search("\d\)", before_italics[-8:]) and re.search("[\d\s()–\-,]*", after_italics[:6]): # identifies the journal article
                # reference_type = "journal article"

                # possible errors
                if before_italics[-1] != " ":
                    error_references.append("Reference " + str(i + 1) + " is missing a space before the journal title.")
                if before_italics[-2] != ".":
                    if before_italics[-2] == "\"":
                        error_references.append("In Reference " + str(i + 1) + ", the article title should not be in quotation marks, as per APA standards.")
                    if before_italics[-2] != ".\"":
                        pass
                    else:
                        error_references.append("Reference " + str(i + 1) + " should have a period after the article title.")

                if after_italics[0] != ",":
                    error_references.append("Reference " + str(i + 1) + " should have a comma after the journal title.")
                if not re.search("\d{1,3}\(\d{1,2}\)", after_italics):
                    error_references.append("In Reference " + str(i + 1) + ", the volume and issue information should conform to the following format: 43(3), with no space before the opening bracket.")
                if "p." in after_italics:
                    error_references.append("In Reference " + str(i + 1) + ", page number should not be preceded by 'p.' or 'pp.'. Page numbering should conform to the following format: '23-56'.")
                if italicized[i].isupper():
                    error_references.append("In Reference " + str(i + 1) + ", the journal title should not be all in uppercase.")

            # website and newspapers
            if "https://" in reference_list[i] and "https://doi" not in reference_list[i]\
                    and reference_list[i].index("https://") > reference_list[i].index(italicized[i]): # and italicized is just before the http
                # reference_type = "website or newspaper"
                # print(reference_list[i])
                pass


        else:
            error_references.append("In Reference " + str(i + 1) + ", the title is not in italics. Remember: either the journal title, book title or website title should always be in italics, even in the body.")


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
# print("Final report: ", final_report)

