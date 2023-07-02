# before running the program you need to run next commands in the terminal
# pip install pdfminer.six
# pip install openpyxl
import re
from pdfminer.high_level import extract_text
from openpyxl import load_workbook
from openpyxl.utils.exceptions import IllegalCharacterError
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE


# function to define the first empty cell in the excel file
def first_empty():
    wb = load_workbook("task.xlsx")
    ws = wb.active
    counter = 3

    while ws[f"A{counter}"].value:
        counter += 1

    return counter


# function to turn the pdf file in the task into a list of separate articles
def pdf_to_chunks(pdf_file: str):
    text: str = extract_text(pdf_file)
    split_text: list = re.split(r"(?=\bP\d{3}\b)", text)
    split_text[-1] = re.split(r"www.medicaljournals.se/acta5th World Psoriasis & Psoriatic Arthritis Conference 2018",
                              split_text[-1])
    split_text[-1] = split_text[-1][0]

    return split_text


# function to work with some specific data in the pdf (the one with several authors with different affiliations)
def return_digit(text: str):
    for i in text:
        if i.isdigit():
            return i

    return False


# function to clean and to normalize text. I suggest it can be improved to refine the results
def clean_text(txt: str):
    cleaned_text: str = re.sub(r'www.medicaljournals.se/acta5th World Psoriasis & Psoriatic Arthritis Conference 2018',
                               '', txt)
    cleaned_text: str = ILLEGAL_CHARACTERS_RE.sub('', cleaned_text)

    return cleaned_text


# function to create different regex pattern to find an article or to return False in case such pattern isn't found
def create_presentation_pattern(text: str):
    possible_beginnings: list = [
        "Introduction:",
        "Background:",
        "Objective:",
        "Introduction/Objective:",
        "Background/Objective:",
        "Introduction & Objectives:",
        "Introduction and Objectives:"
    ]
    beginnings: list = []

    for beginning in possible_beginnings:
        if beginning in text:
            beginnings.append(beginning)

    if not beginnings:
        return False

    return beginnings[0]


# here we create a list of articles and clean them so that the parsing would be more precise
split_text: list = pdf_to_chunks(
    "Abstract Book from the 5th World Psoriasis and Psoriatic Arthritis Conference 2018.pdf")

cleaned_texts: list = [clean_text(text) for text in split_text][1:]


# class that represents each article
class Article:
    # unfortunately I couldn't find a pattern to make affiliation and location different instance attributes
    # so here I decided to merge them
    # but I believe I can solve this later given more time
    def __init__(self, person: list, affiliation_location: list, session: str, topic_title: str, presentation: str):
        self.person = person
        self.affiliation_location = affiliation_location
        self.session = session
        self.topic_title = topic_title
        self.presentation = presentation

    # this method uploads the data from the class into the excel file in the format provided in the task
    def upload_excel(self):
        wb = load_workbook("task.xlsx")
        ws = wb.active
        counter = first_empty()

        for person in self.person:
            ws[f"A{counter + self.person.index(person)}"].value = person

            for affiliation in self.affiliation_location:
                if return_digit(affiliation) == return_digit(person):
                    ws[f"B{counter + self.person.index(person)}"].value = affiliation

            ws[f"D{counter + self.person.index(person)}"].value = self.session
            ws[f"E{counter + self.person.index(person)}"].value = self.topic_title
            ws[f"F{counter + self.person.index(person)}"].value = self.presentation

        wb.save("task.xlsx")


# function that creates an Article instance from a single article from the pdf
def group_article(txt: str):
    pattern_word = create_presentation_pattern(txt)
    lines = txt.split("\n")
    person_lines = [line for line in lines[1:] if line.istitle()]
    topic_title_lines = [line for line in lines[1:] if line.isupper()]
    person = [single_name.lstrip() for single_name in ''.join(person_lines).split(',')]
    session = re.findall(r"\bP\d{3}\b", txt)[0]

    if pattern_word:
        info = re.split(r'(?={})'.format(pattern_word), txt, flags=re.IGNORECASE)
        affiliation_location_lines = [item for item in info[0].split("\n") if
                                      item not in person_lines and item not in topic_title_lines][1:]
        affiliation_location = re.split(r'(?=\d[A-Z])', ''.join(affiliation_location_lines))
        affiliation_location = [item for item in affiliation_location if item != '']
        topic_title = ''.join(topic_title_lines)
        presentation = info[1]
    else:
        affiliation_location = "THE ARTICLE SEEMS TO HAVE NON-STANDARD PATTERN, YOU MIGHT NEED TO ADD IT MANUALLY"
        topic_title = ''.join(topic_title_lines)
        presentation = "THE ARTICLE SEEMS TO HAVE NON-STANDARD PATTERN, YOU MIGHT NEED TO ADD IT MANUALLY"

    return Article(person, affiliation_location, session, topic_title, presentation)


# encountered_errors list and try/except were used to avoid IllegalCharacterError previously
encountered_errors = []
for text in cleaned_texts:
    try:
        article = group_article(text)
        article.upload_excel()
    except IllegalCharacterError as e:
        encountered_errors.append(article.session)
        print(f"We have encountered an Error in {article.session}")

# I'm aware that the results of the parsing could be better and I'm sure I can improve the results
