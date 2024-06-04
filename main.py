import os
import re
from docx import Document

def list_files_in_directory(directory_path):
    return [os.path.join(directory_path, f) for f in os.listdir(directory_path) if os.path.isfile(os.path.join(directory_path, f))]

def read_docx_file(file_path):
    doc = Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

def parse_question(question_text):
    """
    Parse the question text and determine the schema type.
    """
    # Patterns for different question types
    single_choice_pattern = re.compile(r"(\d+\s+\w.*?(\n|$))+")
    numeric_pattern = re.compile(r"\d{4}\s.*?(\n|$)")
    open_ended_pattern = re.compile(r"Please describe.*|RECORD RESPONSE VERBATIM")
    scale_pattern = re.compile(r"Using a scale from \d to \d")
    dropdown_pattern = re.compile(r"\(ALPHABETIZED DROP DOWN LIST.*\)")
    other_dropdown_pattern = re.compile(r"(DROP DOWN LIST OF .+)")
    multiple_choice_pattern = re.compile(r"Please select all that apply|Select all that apply")
    thank_terminate_pattern = re.compile(r"\(THANK AND TERMINATE\)")
    next_pattern = re.compile(r"(MOVING ON|CHANGE TOPIC|NEXT QUESTION|SHOW ON SAME PAGE)")

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

def extract_dropdown_options(question_text):
    """
    Extract dropdown options from the question text.
    """
    options = re.findall(r"(\d+)\s+(.*?)(\n|$)", question_text)
    return "\n".join([f"{int(opt[0]):02d} {opt[1]}" for opt in options])

def generate_markup(question_text, question_label):
    """
    Generate the appropriate markup for the given question text.
    """
    question_type = parse_question(question_text)
    
    if question_type == "single_choice":
        options = re.findall(r"(\d+)\s+(.*?)(\n|$)", question_text)
        options_markup = "\n".join([f"{int(opt[0]):02d} {opt[1]}" for opt in options])
        return f"{{{question_label}:\n{question_text.strip()}\n!FIELD\n{options_markup}\n}}"
    
    elif question_type == "numeric":
        return f"{{{question_label}:\n{question_text.strip()}\n!NUMERIC,,,1900-2024,9999\n}}"
    
    elif question_type == "open_ended":
        return f"{{{question_label}:\n{question_text.strip()}\n!VERBATIM\n}}"
    
    elif question_type == "scale":
        return f"{{{question_label}:\n{question_text.strip()}\n!NUMERIC,,,1-10\n}}"
    
    elif question_type == "dropdown":
        options_markup = extract_dropdown_options(question_text)
        if not options_markup:  # If no specific options are extracted, add a placeholder
            options_markup = "01 Option 1\n02 Option 2\n03 Option 3"
        return f"{{{question_label}:\n{question_text.strip()}\n!DROPDOWN\n{options_markup}\n}}"
    
    elif question_type == "multiple_choice":
        options = re.findall(r"(\d+)\s+(.*?)(\n|$)", question_text)
        options_markup = "\n".join([f"{int(opt[0]):02d} {opt[1]}" for opt in options])
        return f"{{{question_label}:\n{question_text.strip()}\n!MULTI\n{options_markup}\n}}"
    
    elif question_type == "thank_terminate":
        return f"{{{question_label}:\n{question_text.strip()}\n!FIELD\n{question_text.strip()}\n}}"
    
    elif question_type == "next":
        return f"{{{question_label}:\n{question_text.strip()}\n!FIELD\n{question_text.strip()}\n}}"
    
    else:
        return f"{{{question_label}:\n{question_text.strip()}\n!UNKNOWN\n}}"

def write_markup_to_file(file_path, markup_list):
    output_file_path = f"{os.path.splitext(file_path)[0]}_output.txt"
    with open(output_file_path, 'w', encoding='utf-8') as file:
        for markup in markup_list:
            file.write(markup + "\n\n")
    print(f"Generated markup written to: {output_file_path}")

# Specify the directory containing the files
directory_path = '.'  # Change this path if needed

# List all files in the directory
file_paths = list_files_in_directory(directory_path)

# Process each file
for file_path in file_paths:
    if file_path.endswith('.py') or file_path.startswith('~$'):
        # Skip the script file itself and temporary files
        continue

    print(f"Processing file: {file_path}")
    
    try:
        file_content = read_docx_file(file_path)
        print(f"Read content from {file_path}, length: {len(file_content)} characters")

        # Split content into questions and generate markup
        questions = re.split(r'\n\s*\n', file_content)  # Assuming questions are separated by blank lines
        markup_list = []

        for i, question_text in enumerate(questions):
            question_label = f"Q{i+1}"
            markup = generate_markup(question_text, question_label)
            markup_list.append(markup)
        
        # Write generated markup to file
        write_markup_to_file(file_path, markup_list)
    
    except Exception as e:
        print(f"Error processing file {file_path}: {e}")

print("Processing complete.")
