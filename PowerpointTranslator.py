from pptx import Presentation
import six
import os
from google.cloud import translate_v2 as translate
from pptx.dml.color import RGBColor
import copy
import sys

def copy_font(input_run):
    font_copy = []
    try:
        font_copy = copy.deepcopy(input_run.runs[0].font)
    except Exception as e:
        print("Deepcopy exception ", e.__class__, " occurred.")
    return font_copy

def apply_font(target, saved_font):
    try:
        target.font.bold = saved_font.bold
    except Exception as e:
        print("Bold exception ", e.__class__, " occurred.")

    try:
        target.font.color.rgb = RGBColor(saved_font.color.rgb[0],
                                         saved_font.color.rgb[1],
                                         saved_font.color.rgb[2])
    except Exception as e:
        print("RGB exception ", e.__class__, " occurred.")

    try:
        target.font.italic = saved_font.italic
    except Exception as e:
        print("Italics exception ", e.__class__, " occurred.")

    try:
        target.font.name = saved_font.name
    except Exception as e:
        print("Font name exception ", e.__class__, " occurred.")

    try:
        target.font.size = saved_font.size
    except Exception as e:
        print("Font size exception ", e.__class__, " occurred.")

    try:
        target.font.underline = saved_font.underline
    except Exception as e:
        print("Underline exception ", e.__class__, " occurred.")
    return


def translate_text(translate_client, target, text):

    if isinstance(text, six.binary_type):
        text = text.decode("utf-8")

    # Passing incorrect target language codes will throw unexpected 403 missing a valid API key error.
    # Translation API sometimes detects 'es' instead 'pt', passing source_language improves translation quality.
    result = translate_client.translate(text, target_language=target, source_language='pt', format_='text')

    return result


os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = r"C:\Users\DFLO\PycharmProjects\My_GCP_Translation_API_Key.json"
translate_client = translate.Client()

filename = sys.argv[1]
prs = Presentation()
prs = Presentation(filename)

total_slide_cnt = len(prs.slides)
print(len(prs.slides))

for slide in prs.slides:
    print("slide " + str(prs.slides.index(slide) + 1) + "/" + str(total_slide_cnt))
    for shape in slide.shapes:
        if (shape.has_text_frame):
            for paragraph in shape.text_frame.paragraphs:
                if(paragraph.text != ''):
                    result = translate_text(translate_client, 'en', paragraph.text)

                    saved_font = copy_font(paragraph)

                    # Clears all font settings, so must reapply them.
                    paragraph.text = result["translatedText"]

                    apply_font(paragraph.runs[0], saved_font)

        elif(shape.has_table):
            for cell in shape.table.iter_cells():
                if not cell.is_spanned and cell.text != '':

                    result = translate_text(translate_client, 'en', cell.text)

                    saved_font = copy_font(cell.text_frame.paragraphs[0])

                    # Clears all font settings, so must reapply them.
                    cell.text = result["translatedText"]

                    apply_font(cell.text_frame.paragraphs[0], saved_font)

output_filename = filename.split('.')[0] + '_en2.pptx'
prs.save(output_filename)