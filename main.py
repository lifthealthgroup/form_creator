import pandas as pd
import fitz
from flask import Flask, request, send_file, render_template
import os

app = Flask(__name__)

def render_to_image(filled_form, zoom=2):
    """
    Renders the filled PDF form to images and saves them as new PDFs.
    """
    temp_pdf = fitz.open()  # Create a new PDF
    for page_number in range(len(filled_form)):
        page = filled_form[page_number]
        # Render page to an image
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))  # Zoom for better quality
        img_pdf = fitz.open()  # New PDF for this page
        img_page = img_pdf.new_page(width=pix.width, height=pix.height)  # Create a new page
        img_page.insert_image(img_page.rect, stream=pix.tobytes())  # Insert image into the new page
        temp_pdf.insert_pdf(img_pdf)  # Insert the image PDF into the temp PDF

    return temp_pdf

def read_excel(excel:str):
    """
    Reads in path to excel file and populates relevant dictionaries with values.
    Stores dictionaries in master dictionary and returns master dictionary. 
    """

    df = pd.read_excel(excel) # read excel into pandas df
    master = {} # initialise master dictionary
    for col in df.columns:
        if not pd.isna(df[col].iloc[0]) and 'values' in col.lower(): # find dictionaries to populate
            dict_name = col.replace(' Values', '') # dictionary key name
            master[dict_name] = pd.Series(df[col].values, index=df[dict_name]).to_dict() # place dictionary in dictionary
    return master # which contains all info needed to fill forms

def fill_textboxes(general_values:dict, form_values:dict, template):
    """
    Fills textbox values in pdf based on values in dictionary.
    """
    for page in template:
        for field in page.widgets(): # iterate through fields on each page
            key = field.field_name
            if key in form_values: # add form values
                field.field_value = str(form_values[key])
                field.update()
            if key in general_values: # add general values
                field.field_value = str(general_values[key])
                field.update()
            if key == 'Y1':
                field.field_value = True 
                field.update()
    return template

def highlight(x0, y0, x1, y1, page):
    """
    Highlights the area on template defined by coords x0, y0, x1, y1
    """
    rect = fitz.Rect(x0, y0, x1, y1) # create rectangle object
    highlight = page.add_highlight_annot(rect) # Add a highlight annotation to the defined box
    highlight.update() # save changes
    return page
    
def fill_WHODAS(general_values:dict, form_values:dict):
    """
    Inserts values from dictionary in correct fields in WHODAS pdf file.
    """
    template = fitz.open('forms/WHODAS.pdf')

    # Calculate Extra Fields
    for i in range(1,7): 
        if i != 5:
            number_params = sum(1 for key, _ in form_values.items() if key.startswith('D' + str(i)))
            form_values[str(i) + '_overall'] = sum(value for key, value in form_values.items() if key.startswith('D' + str(i)))
            form_values[str(i) + '_avg'] = round(sum(value for key, value in form_values.items() if key.startswith('D' + str(i)))/number_params , 2)
    form_values['5_overall'] = form_values['D51'] + form_values['D52'] + form_values['D53'] + form_values['D51']
    form_values['5_avg'] = round(form_values['5_overall'] / 4, 2)
    form_values['5_overall2'] = form_values['D55'] + form_values['D56'] + form_values['D57'] + form_values['D58']
    form_values['5_avg2'] = round(form_values['5_overall2'] / 4, 2)
    form_values['total'] = form_values['1_overall'] + form_values['2_overall'] + form_values['3_overall'] + form_values['4_overall'] + form_values['5_overall'] + form_values['5_overall2'] + form_values['6_overall']
    form_values['avg'] = round(form_values['total'] / 36, 2)
    
    # Fill in textboxes
    template = fill_textboxes(general_values, form_values, template)

    # Fill in checkboxes
    for page in template:
        for field in page.widgets():
            key = field.field_name
            # Check if it's a checkbox
            if field.field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
                if key == 'male' and general_values['gender'].lower() == 'm':
                    field.field_value = True  # Set checkbox to checked
                elif key == 'female' and general_values['gender'].lower() == 'f':
                    field.field_value = True  # Set checkbox to checked
                field.update()
    return template

def fill_CANS(general_values:dict, form_values:dict):
    """
    Inserts values from dictionary in correct fields in CANS pdf file.
    """
    
    template = fitz.open('forms/CANS.pdf') # read in template pdf

    form_values['A_subtotal'] = 0
    form_values['B_subtotal'] = 0
    form_values['C_subtotal'] = 0
    form_values['D_subtotal'] = 0
    for key in form_values.keys():
        try:
            if form_values[key].upper() == 'Y':
                if key > 0 and key < 11:
                    form_values['A_subtotal'] += 1
                elif key > 10 and key < 15:
                    form_values['B_subtotal'] += 1
                elif key > 14 and key < 26:
                    form_values['C_subtotal'] += 1
                elif key > 25 and key < 29:
                    form_values['D_subtotal'] += 1
        except Exception:
            continue

    form_values['subtotal'] = form_values['A_subtotal'] + form_values['B_subtotal'] + form_values['C_subtotal'] + form_values['D_subtotal']

    # calculate CANS level
    if form_values['A_subtotal'] < 4:
        if form_values['B_subtotal'] >= 4:
            form_values['total'] = 4.2
        elif form_values['C_subtotal'] >= 4:
            form_values['total'] = 4.1
        elif form_values['C_subtotal'] == 3:
            form_values['total'] = 3
        elif form_values['C_subtotal'] == 2:
            form_values['total'] = 2    
        elif form_values['C_subtotal'] == 1:
            form_values['total'] = 1
        else:
            form_values['total'] = 0    
    elif form_values['A_subtotal'] == 5:
        form_values['total'] = 5
    elif form_values['A_subtotal'] == 6:
        form_values['total'] = 6
    else:
        form_values['total'] = 7
    
    template = fill_textboxes(general_values, form_values, template) # fill out textboxes

    for page in template: # gather fields
        for field in page.widgets():
            key = field.field_name
            # Check if it's a checkbox
            if field.field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
                if form_values[int(key[1:])].upper() == 'Y' and key[:1] == 'Y':
                    field.field_value = True  # Set checkbox to checked
                elif form_values[int(key[1:])].upper() == 'N' and key[:1] == 'N':
                    field.field_value = True  # Set checkbox to checked
                field.update()

    return template

def fill_LSP(general_values:dict, form_values:dict):
    """
    Inserts values from dictionary in correct fields in LSP pdf file.
    """
    template = fitz.open('forms/LSP.pdf') # read in template pdf
    
    # define coordinates of grid layout on LSP page 1
    x = [550, 670, 792, 912, 1034]
    x = [px * (72 / 140) for px in x] # convert pixels to coordinates. dpi = 140
    y = [328,380,454,506,580,652,744,778,892,944,1038,1090,1164,1236,1288,1320,1394]
    y = [px * (72 / 140) for px in y]

    # highlight correct box for each row
    page = template.load_page(0)
    for i in range(16):
        score = form_values[i + 1]
        if score == 0:
            page = highlight(x[0] + 10, y[i], x[1] - 10, y[i + 1], page) # bring in the highlight slightly due to formatting
        elif score == 1:
            page = highlight(x[1] + 10, y[i], x[2] - 10, y[i + 1], page) # bring in the highlight slightly due to formatting
        elif score == 2:
            page = highlight(x[2] + 10, y[i], x[3] - 10, y[i + 1], page) # bring in the highlight slightly due to formatting
        elif score == 3:
            page = highlight(x[3] + 10, y[i], x[4] - 10, y[i + 1], page) # bring in the highlight slightly due to formatting
    
    # perform scoring
    form_values['a_score'] = int(form_values[1] + form_values[2] + form_values[3] + form_values[8])
    form_values['b_score'] = int(form_values[4] + form_values[5] + form_values[6] + form_values[9] + form_values[16])
    form_values['c_score'] = int(form_values[10] + form_values[11] + form_values[12])
    form_values['d_score'] = int(form_values[7] + form_values[13] + form_values[14] + form_values[15])
    
    total = form_values['a_score'] + form_values['b_score'] + form_values['c_score'] + form_values['d_score']
    form_values['total'] = "Total Score:  " + str(total) + "/80 = " + str(total * 1.25) + "/100"
    
    form_values['a_score'] = str(form_values['a_score']) + "/12"
    form_values['b_score'] = str(form_values['b_score']) + "/15"
    form_values['c_score'] = str(form_values['c_score']) + "/9" 
    form_values['d_score'] = str(form_values['d_score']) + "/12"
    
    template = fill_textboxes(general_values, form_values, template) # fill out textboxes

    return template

def produce_output(master:dict[dict]):
    """
    Calls fill_form function for each dictionary in dictionaries and combines pdfs to final file. 
    """
    combined = fitz.open() # new document
    for key in master.keys():
        if key != 'GENERAL':
            function_name = globals().get(f"fill_{key}")
            if function_name:
                filled_form = function_name(master['GENERAL'], master[key])
                rendered_pdf = render_to_image(filled_form) # this is a workaround to fuse field values to page 
                combined.insert_pdf(rendered_pdf) # append to combined
    return combined

@app.route('/')
def index():
    return render_template('upload.html')  # Simple upload form

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part", 400
    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400

    # Process the uploaded file
    master = read_excel(file.stream)  # Read Excel from the uploaded file
    final_document = produce_output(master)
    
    # Save the output PDF
    output_path = 'output/new.pdf'
    if not os.path.exists('output'):
        os.makedirs('output')
    
    final_document.save(output_path)
    
    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)