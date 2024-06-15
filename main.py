import os
import re
from flask import Flask, request, redirect, url_for, render_template, send_from_directory
from werkzeug.utils import secure_filename
from docx import Document
from openai import OpenAI
from concurrent.futures import ProcessPoolExecutor

client = OpenAI(api_key='sk-proj-2rbgPrswTlqiyVVaDAMHT3BlbkFJ8HXX5Wx9nAHSdcZpZsTC')

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'docx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def read_docx_file(file_path):
    doc = Document(file_path)
    full_text = [para.text for para in doc.paragraphs]
    return '\n'.join(full_text)

def process_file(file_path):
    try:
        content = read_docx_file(file_path)
        questions = split_questions(content)
        questions = merge_related_content(questions)
        markup_list = [generate_markup(question, f"Q{i+1}") for i, question in enumerate(questions)]
        refined_markup = call_openai_api(markup_list)
        output_filename = f"{os.path.splitext(os.path.basename(file_path))[0]}_output.txt"
        output_file_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        with open(output_file_path, 'w', encoding='utf-8') as file:
            for markup in refined_markup:
                file.write(markup + "\n\n")
        return output_filename
    except Exception as e:
        print(f"Error processing file {file_path}: {e}")
        raise

def process_files(file_paths):
    try:
        with ProcessPoolExecutor() as executor:
            results = list(executor.map(process_file, file_paths))
        return results
    except Exception as e:
        print(f"Error processing files: {e}")
        raise

@app.route('/', methods=['GET', 'POST'])
def upload_files():
    if request.method == 'POST':
        if 'files[]' not in request.files:
            return redirect(request.url)
        files = request.files.getlist('files[]')
        file_paths = []

        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(file_path)
                file_paths.append(file_path)

        if file_paths:
            try:
                output_files = process_files(file_paths)
                return render_template('upload.html', files=output_files)
            except Exception as e:
                return f"An error occurred while processing the files: {e}"

    return render_template('upload.html')

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(debug=True)

# Your existing functions
def split_questions(content):
    return re.split(r'(?m)^\s*(?=[A-Z]\d*\.|\d+\s)', content)

def merge_related_content(questions):
    merged_questions = []
    buffer = []

    for question in questions:
        if re.match(r'^[A-Z]\d*\.', question) and buffer:
            merged_questions.append('\n'.join(buffer))
            buffer = []
        buffer.append(question.strip())

    if buffer:
        merged_questions.append('\n'.join(buffer))

    return merged_questions

def parse_question(question_text):
    single_choice_pattern = re.compile(r"^\d+\s+.*$", re.MULTILINE)
    numeric_pattern = re.compile(r"\d{4}\s.*?\d{4}", re.MULTILINE)
    open_ended_pattern = re.compile(r"Please describe.*|RECORD RESPONSE VERBATIM", re.MULTILINE)
    scale_pattern = re.compile(r"Using a scale from \d to \d", re.MULTILINE)
    dropdown_pattern = re.compile(r"\(ALPHABETIZED DROP DOWN LIST.*\)", re.MULTILINE)
    other_dropdown_pattern = re.compile(r"(DROP DOWN LIST OF .+)", re.MULTILINE)
    multiple_choice_pattern = re.compile(r"Please select all that apply|Select all that apply", re.MULTILINE)
    thank_terminate_pattern = re.compile(r"\(THANK AND TERMINATE\)", re.MULTILINE)
    next_pattern = re.compile(r"(MOVING ON|CHANGE TOPIC|NEXT QUESTION|SHOW ON SAME PAGE)", re.MULTILINE)

    if single_choice_pattern.search(question_text):
        return "single_choice"
    elif numeric_pattern.search(question_text):
        return "numeric"
    elif open_ended_pattern.search(question_text):
        return "open_ended"
    elif scale_pattern.search(question_text):
        return "scale"
    elif dropdown_pattern.search(question_text) or other_dropdown_pattern.search(question_text):
        return "dropdown"
    elif multiple_choice_pattern.search(question_text):
        return "multiple_choice"
    elif thank_terminate_pattern.search(question_text):
        return "thank_terminate"
    elif next_pattern.search(question_text):
        return "next"
    else:
        return "unknown"

def extract_options(question_text):
    options = re.findall(r"^\d+\s+.*$", question_text, re.MULTILINE)
    return "\n".join([f"{int(opt.split()[0]):02d} {' '.join(opt.split()[1:])}" for opt in options])

def generate_markup(question_text, question_label):
    question_type = parse_question(question_text)

    if question_type == "single_choice":
        options_markup = extract_options(question_text)
        return f"{{{question_label}:\n{question_text.strip()}\n!FIELD\n{options_markup}\n}}"

    elif question_type == "numeric":
        return f"{{{question_label}:\n{question_text.strip()}\n!NUMERIC,,,1900-2024,9999\n}}"

    elif question_type == "open_ended":
        return f"{{{question_label}:\n{question_text.strip()}\n!VERBATIM\n}}"

    elif question_type == "scale":
        return f"{{{question_label}:\n{question_text.strip()}\n!NUMERIC,,,1-10\n}}"

    elif question_type == "dropdown":
        options_markup = extract_options(question_text)
        if not options_markup:
            options_markup = "01 Option 1\n02 Option 2\n03 Option 3"
        return f"{{{question_label}:\n{question_text.strip()}\n!DROPDOWN\n{options_markup}\n}}"

    elif question_type == "multiple_choice":
        options_markup = extract_options(question_text)
        return f"{{{question_label}:\n{question_text.strip()}\n!MULTI\n{options_markup}\n}}"

    elif question_type == "thank_terminate":
        return f"{{{question_label}:\n{question_text.strip()}\n!FIELD\n{question_text.strip()}\n}}"

    elif question_type == "next":
        return f"{{{question_label}:\n{question_text.strip()}\n!FIELD\n{question_text.strip()}\n}}"

    else:
        return f"{{{question_label}:\n{question_text.strip()}\n!UNKNOWN\n}}"

def call_openai_api(markup_list):
    instructions = """
    You are refining a questionnaire for survey markup. Each question should follow this structure:
    - Start with a left curly bracket '{'
    - Followed by the Question Label, which starts with an uppercase 'Q'
    - The Question Label is immediately followed by a colon ':'
    - The following lines contain the question text
    - The 'Question Type' line terminates the question text lines. It starts with an exclamation point '!' followed by the question type.
    - For single choice questions, use '!FIELD' followed by the answer choices/categories. Each line starts with the answer code and is followed by the answer text. Answer codes must be numeric and zero-padded.
    - For numeric questions, use '!NUMERIC,,,min,max,refused_code'
    - For open-ended questions, use '!VERBATIM'
    - For dropdown questions, use '!DROPDOWN' followed by the options
    - The question definition is terminated with a right curly bracket '}'

    Here is an example question:
    {Q1:
    Is the country going in the right direction or is it on the wrong track?
    !FIELD
    01 RIGHT DIRECTION
    02 WRONG TRACK
    08 DON’T KNOW
    09 REFUSED
    }

    Now refine the following questionnaire markup:
    """

    refined_markup = []
    for markup in markup_list:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a survey markup expert."},
                {"role": "user", "content": f"{instructions}\n\n{markup}"}
            ]
        )
        refined_markup.append(response.choices[0].message.content.strip())
    return refined_markup
