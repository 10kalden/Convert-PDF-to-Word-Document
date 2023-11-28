import os
from docx import Document
from transformers import pipeline

def load_model(model_name):
    try:
        return pipeline('text-generation', model=model_name)
    except Exception as e:
        print(f"Error loading model: {e}")
        return None

def open_document(file_path):
    if not os.path.exists(file_path):
        print(f"File does not exist: {file_path}")
        return None

    try:
        return Document(file_path)
    except Exception as e:
        print(f"Error opening document: {e}")
        return None

def generate_text(generator, prompt, max_length):
    try:
        prompt_length = len(prompt.split())
        max_length = max(max_length, prompt_length + 1)
        if len(prompt) > max_length:
            print("Prompt is longer than max_length, truncating...")
            prompt = prompt[:max_length]
        return generator(prompt, max_length=max_length)[0]['generated_text']
    except Exception as e:
        print(f"Error generating text: {e}")
        return None

def save_document(doc, file_path):
    try:
        doc.save(file_path)
    except Exception as e:
        print(f"Error saving document: {e}")

def main():
    generator = load_model('gpt2')
    if generator is None:
        return

    doc = open_document('new.docx') #input docs
    if doc is None:
        return

    title = doc.paragraphs[0].text if doc.paragraphs else ""

    generated = generate_text(generator, title, 500)
    if generated is None:
        return

    content = generated[len(title):]

    doc.add_paragraph(content)

    save_document(doc, "updated_doc.docx") #output doc

if __name__ == "__main__":
    main()
