import json

from flask import Flask, request, jsonify
from docx import Document
import convertapi
import fitz
from flask_cors import CORS, cross_origin

app = Flask(__name__)
CORS(app, origins="*")

inputFile = "myfolder/document.docx"
outputFile = "myfolder/document.pdf"


@app.route('/', methods=['GET'])
@cross_origin()
def hello():
    return "Hello World"


@app.route('/check-text/', methods=['POST'])
@cross_origin()
def hello_world():
    file = request.files['file']
    paramsText = json.loads(request.form.get('data'))

    # Нужно раскомитить
    file.save(inputFile)
    convertapi.api_secret = 'G3LHidhmyZOOYX56'
    result = convertapi.convert('pdf', {'File': inputFile})
    result.file.save(outputFile)

    try:
        return jsonify(checkDocument(file, paramsText))
    except ValueError:
        return ValueError


def checkDocument(dataFile, params):
    atr = []
    document = Document(dataFile)
    resultTest = {}
    result = []

    for index, paragraph in enumerate(document.paragraphs):
        for param in params:
            if param == 'fontSize':
                atr = atr + getFontSizeForParagraph(paragraph, document)
            if param == 'fontName':
                atr = atr + getFontsForParagraph(paragraph, document)
            if param == 'firstLineIndent':
                atr = atr + getFirstLineIndentForParagraph(paragraph, document)

            unique = list(map(lambda x: str(x), list(set(atr))))
            if unique != [params.get(param)]:
                if param in resultTest:
                    resultTest[param] = resultTest[param] + [getPageByParagraph(paragraph.text)]
                else:
                    resultTest[param] = [getPageByParagraph(paragraph.text)]
            atr = []

    for res in resultTest:
        result.append({'name': res, 'listProblemsPage': list(set(resultTest[res]))})

    return result


def getPageByParagraph(paragraph):
    rim = fitz.open(outputFile)
    for current_page in range(len(rim)):
        page = rim.load_page(current_page)
        if page.search_for(paragraph):
            return current_page + 1


def getFontSizeForParagraph(paragraph, document):
    result = []
    for run in paragraph.runs:
        fontSizeParagraph = run.font.size
        fontSizeFromStyles = document.styles[getStyleParagraph(paragraph)].font.size
        if fontSizeParagraph == 0:
            result.append(round(0))
        elif fontSizeParagraph:
            result.append(round(fontSizeParagraph.pt))
        else:
            result.append(round(fontSizeFromStyles.pt))
    return result


def getFirstLineIndentForParagraph(paragraph, document):
    lineIndentParagraph = paragraph.paragraph_format.first_line_indent
    lineIndentFromStyles = document.styles[getStyleParagraph(paragraph)].paragraph_format.first_line_indent
    if lineIndentParagraph == 0:
        return [round(0, 2)]
    elif lineIndentParagraph:
        return [round(lineIndentParagraph.cm, 2)]
    return [round(lineIndentFromStyles.cm, 2)]


def getFontsForParagraph(paragraph, document):
    result = []
    for run in paragraph.runs:
        result.append(run.font.name or document.styles[getStyleParagraph(paragraph)].font.name)
    return result


def getStyleParagraph(paragraph): return paragraph.style.name


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')
