import xlrd
import requests

FILE_DIR = "QA_new.xls"
access_token = 'NjkwLTItMTAwOTQtNDZmNDhhNGZhZmRlNjc4MGMyNDJiYzM4YzU2Mzk1YzYzMWY3NjkwYzM1NjM4ODAwZWUzZGNhMjMxOGYzMjc3NA=='
url = 'https://appadmin.mipt.ru/api/bot-answer/add/'
headers = {
    "Authorization": "Bearer " + access_token,
}


class ButtonItem:
    def __init__(self, name="", answer=""):
        self.name = name
        self.answer = answer


class QuestionItem:
    def __init__(self, name="", keyword="", answer="", buttons=None):
        if buttons is None:
            buttons = []
        self.name = name
        self.keyword = keyword
        self.answer = answer
        self.buttons = buttons


def parse_book(path):
    book = xlrd.open_workbook(path)
    # get the first worksheet
    sheet = book.sheet_by_index(0)

    questions = []
    for row in range(1, sheet.nrows):
        question = sheet.cell(colx=0, rowx=row).value
        keyword = sheet.cell(colx=1, rowx=row).value.replace(',', ' ')
        answer = sheet.cell(colx=2, rowx=row).value
        button_title = sheet.cell(colx=3, rowx=row).value
        button_answer = sheet.cell(colx=4, rowx=row).value
        if question != '':
            item = QuestionItem(name=question, keyword=keyword, answer=answer)
            if button_title != '':
                button = ButtonItem(name=button_title, answer=button_answer)
                item.buttons.append(button)
            questions.append(item)
        elif len(questions) != 0:
            questions[-1].buttons.append(ButtonItem(name=button_title, answer=button_answer))

    return questions


def upload_question(question):
    data = {
        "bot-answer-form[bot-answer][name]": question.name,
        "bot-answer-form[bot-answer][keyword]": question.keyword,
        "bot-answer-form[bot-answer][answer]": question.answer,
    }
    for i in range(len(question.buttons)):
        data['bot-answer-form[bot-answer][buttons][%d][name]' % i] = question.buttons[i].name
        data['bot-answer-form[bot-answer][buttons][%d][answer]' % i] = question.buttons[i].answer

    try:
        response = requests.post(url, headers=headers, data=data)
        # print(response.json())
    except Exception as e:
        print(e)
        return -1

    return 0


def main():
    questions = parse_book(FILE_DIR)

    for question in questions:
        result = upload_question(question)
        if result != 0:
            print('Error while uploading question %s' % question.name)


if __name__ == '__main__':
    main()
