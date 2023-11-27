from docx import Document
from transformers import GPT2LMHeadModel, GPT2Tokenizer

# Load pre-trained GPT-2 model and tokenizer
tokenizer = GPT2Tokenizer.from_pretrained("gpt2")
model = GPT2LMHeadModel.from_pretrained("gpt2")

# Load the Word document
doc = Document('output.docx')  # Replace 'your_document.docx' with your file path

# Count the number of lines in the document
line_count = sum(1 for _ in doc.paragraphs)

if line_count < 10:
    # Calculate the number of additional lines needed to reach 20 lines
    additional_lines_needed = 20 - line_count

    # Prompt for content generation
    prompt_text = "\n".join([para.text for para in doc.paragraphs])

    # Generate text to reach the required number of lines
    while additional_lines_needed > 0:
        # Generate text
        input_ids = tokenizer.encode(prompt_text, return_tensors='pt')
        generated_text = model.generate(input_ids, max_length=50, num_return_sequences=1, temperature=0.7)

        # Decode and append the generated text
        generated_text_decoded = tokenizer.decode(generated_text[0], skip_special_tokens=True)
        prompt_text += "\n" + generated_text_decoded

        # Update line count
        additional_lines_needed -= 1

    # Add generated content to the document
    doc.add_paragraph(prompt_text)

    # Save the updated document
    doc.save('extender_doc.docx')  # Save to a new file or overwrite the existing one
