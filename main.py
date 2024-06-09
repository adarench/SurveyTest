import os
from docx import Document
from openai import OpenAI
from concurrent.futures import ThreadPoolExecutor, as_completed

# Initialize the OpenAI client
client = OpenAI(api_key='sk-proj-UjXb6sDuQ1PO1yjtaKJMT3BlbkFJhZEP9SNEEZtioNYiGicK')

def list_files_in_directory(directory_path):
    return [os.path.join(directory_path, f) for f in os.listdir(directory_path) if os.path.isfile(os.path.join(directory_path, f))]

def read_docx_file(file_path):
    doc = Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

def call_openai_api_for_markup(content, file_name):
    instructions = f"""
    You are refining a questionnaire for survey markup. Each question should follow this structure:
    - Start with a left curly bracket '{{'
    - Followed by the Question Label, which starts with an uppercase 'Q'
    - The Question Label is immediately followed by a colon ':'
    - The following lines contain the question text
    - The 'Question Type' line terminates the question text lines. It starts with an exclamation point '!' followed by the question type.
    - For single choice questions, use '!FIELD' followed by the answer choices/categories. Each line starts with the answer code and is followed by the answer text. Answer codes must be numeric and zero-padded.
    - For numeric questions, use '!NUMERIC,,,min,max,refused_code'
    - For open-ended questions, use '!VERBATIM'
    - For dropdown questions, use '!DROPDOWN' followed by the options
    - The question definition is terminated with a right curly bracket '}}'

    Here is an example question:
    {{Q1:
    Is the country going in the right direction or is it on the wrong track?
    !FIELD
    01 RIGHT DIRECTION
    02 WRONG TRACK
    08 DON’T KNOW
    09 REFUSED
    }}

    Now refine the following questionnaire markup for the given file '{file_name}':
    {content}
    """

    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content": "You are a survey markup expert."},
            {"role": "user", "content": instructions}
        ]
    )

    return response.choices[0].message.content.strip()

def write_markup_to_file(file_path, markup):
    output_file_path = f"{os.path.splitext(file_path)[0]}_output.txt"
    with open(output_file_path, 'w', encoding='utf-8') as file:
        file.write(markup)
    print(f"Generated markup written to: {output_file_path}")

def process_file(file_path):
    try:
        file_content = read_docx_file(file_path)
        print(f"Read content from {file_path}, length: {len(file_content)} characters")

        # Refine markup using OpenAI's API
        refined_markup = call_openai_api_for_markup(file_content, os.path.basename(file_path))
        
        # Write refined markup to file
        write_markup_to_file(file_path, refined_markup)
    
    except Exception as e:
        print(f"Error processing file {file_path}: {e}")

# Specify the directory containing the files
directory_path = '.'  # Change this path if needed

# List all files in the directory
file_paths = list_files_in_directory(directory_path)

# Process files in parallel
with ThreadPoolExecutor(max_workers=5) as executor:
    futures = {executor.submit(process_file, file_path): file_path for file_path in file_paths if not file_path.endswith('.py') and not file_path.startswith('~$')}
    for future in as_completed(futures):
        file_path = futures[future]
        try:
            future.result()
        except Exception as exc:
            print(f'{file_path} generated an exception: {exc}')

print("Processing complete.")
