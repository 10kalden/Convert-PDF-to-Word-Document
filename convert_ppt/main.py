import os
import time
import win32com.client
import re

# give full path of ppt
PRESENTATION_PATH = r"C:\Users\tenka\bdc\python\Project_01\convert_ppt\Unit1.pptx" 

# text clean-up func
def sanitize_text(text):
    sanitized_text = re.sub(r'\W+', ' ', text)
    return sanitized_text

# func to extract title and body
def extract_title_body(slide):
    title_shape = next((shape for shape in slide.Shapes if shape.Name.startswith("Title")), None)
    title = title_shape.TextFrame.TextRange.Text if title_shape and title_shape.HasTextFrame else ""
    body = "\n".join(shape.TextFrame.TextRange.Text for shape in slide.Shapes if shape.HasTextFrame and shape != title_shape)
    return sanitize_text(title.strip()), sanitize_text(body.strip())

# func to insert text in word doc
def add_text_range(document, text, bold=False, size=12):
    range = document.Content
    range.Collapse(0) 
    range.InsertAfter(text)
    range.Font.Bold = bold
    range.Font.Size = size

# for displaying in the word doc
def ppt_to_png_and_display_in_word():
    try:
        Application = win32com.client.Dispatch("PowerPoint.Application")
        Presentation = Application.Presentations.Open(PRESENTATION_PATH, WithWindow=False)

        presentation_directory = os.path.dirname(PRESENTATION_PATH)

        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False #hiding the word doc

        word_doc = word_app.Documents.Add()

        slides_folder = os.path.join(presentation_directory, "slide_images") #folder of saved images
        os.makedirs(slides_folder, exist_ok=True)

        for i, slide in enumerate(Presentation.Slides, start=1):
            title, body = extract_title_body(slide)

            image_path = os.path.join(slides_folder, f"{i}.png")
            slide.Export(image_path, "PNG")  # Export the slide as PNG

            # Insert title into the Word document
            if title:
                add_text_range(word_doc, f"{title}\n", bold=True, size=14)

            # Insert the image into the Word document
            word_range = word_doc.Content
            word_range.Collapse(0)  # Collapse to the end of the document
            word_range.InlineShapes.AddPicture(image_path)

            # Insert body text into the Word document
            if body:
                add_text_range(word_doc, f"\n{body}")

            # Insert a page break after each slide
            word_range = word_doc.Content
            word_range.Collapse(0) 
            word_range.InsertBreak(7)  

        # Save the Word doc
        word_doc.SaveAs(os.path.join(presentation_directory, "output.docx"), FileFormat=12) 

        time.sleep(2) 

        Presentation.Close()
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        Application.Quit()
        word_app.Quit()

ppt_to_png_and_display_in_word()
