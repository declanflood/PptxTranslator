from pptx import Presentation
import six
import os
from google.cloud import translate_v2 as translate
import sys


# Translate a string
def translate_text(translate_client, target, text):

    if isinstance(text, six.binary_type):
        text = text.decode("utf-8")

    # Passing incorrect target language codes will throw unexpected 403 missing a valid API key error.
    # Translation API sometimes detects 'es' instead 'pt', passing source_language improves translation quality.
    result = translate_client.translate(text, target_language=target, source_language='pt', format_='text')

    return result


# Translate a paragraph
def translate_paragraph(paragraph):

    if (paragraph.text != ''):
        result = translate_text(translate_client, 'en', paragraph.text)

        p = paragraph._p
        # remove all but the first run (ensures font is not lost)
        for idx, run in enumerate(paragraph.runs):
            if idx == 0:
                continue
            p.remove(run._r)

        if (len(paragraph.runs) == 0):
            paragraph.add_run()

        paragraph.runs[0].text = result["translatedText"]

    return


# Translate all text in a shape
def convert_shape(shape):
    if (shape.has_text_frame):
        # A shape which contains some text
        for paragraph in shape.text_frame.paragraphs:
            translate_paragraph(paragraph)

    elif(shape.has_table):
        # A Table
        for cell in shape.table.iter_cells():
            if not cell.is_spanned and cell.text != '':
                # In tables with merged cells, is_spanned checks if cell visible to user

                for paragraph in cell.text_frame.paragraphs:

                    translate_paragraph(paragraph)

    elif(shape.shape_type == 6):
        # Grouped shapes - use recursion
        for shp in shape.shapes:
            convert_shape(shp)

    return

os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = r".\My_GCP_Translation_API_Key.json"
translate_client = translate.Client()

filename = sys.argv[1]

prs = Presentation(filename)

total_slide_cnt = len(prs.slides)

for slide in prs.slides:
    print("slide " + str(prs.slides.index(slide) + 1) + "/" + str(total_slide_cnt))
    for shape in slide.shapes:
        convert_shape(shape)

output_filename = filename.split('.')[0] + '_en2.pptx'
prs.save(output_filename)
