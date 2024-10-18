import pandas as pd
import fitz
import zipfile
from flask import Flask, request, send_file, render_template, jsonify
import math
from werkzeug.utils import secure_filename
import io
from datetime import datetime

app = Flask(__name__)

def validate_columns(master, file):
    """Validate that if any row in each _Values dictionary is filled, 
       all previous rows must also be filled (but allow trailing empty entries)."""
    errors = []

    # Iterate through each dictionary in the master dictionary
    for dict_name, inner_dict in master.items():
        
        whodas_empties = [['D55', 'D56', 'D57', 'D58'], [55, 56, 57, 58, 59]]
        if dict_name == 'WHODAS':
            # The following keys are considered edge cases, as they can be empty in some circumstances
            has_filled = any(not pd.isna(inner_dict[key]) for key in whodas_empties[0] if key in inner_dict)
            has_empty = any(pd.isna(inner_dict[key]) for key in whodas_empties[0] if key in inner_dict)
            
            # If there are filled and empty entries, validate each key in `whodas_empties`
            if has_filled and has_empty:
                for key in whodas_empties[0]:
                    if key in inner_dict and pd.isna(inner_dict[key]):
                        errors.append(f"In column '{dict_name}', in file '{file}', the field for '{key}' is empty (but others are filled)")
                        
        if dict_name == 'WHODASKIDS':
            # The following keys are considered edge cases, as they can be empty in some circumstances
            has_filled = any(not pd.isna(inner_dict[key]) for key in whodas_empties[1] if key in inner_dict)
            has_empty = any(pd.isna(inner_dict[key]) for key in whodas_empties[1] if key in inner_dict)
            
            # If there are filled and empty entries, validate each key in `whodas_empties`
            if has_filled and has_empty:
                for key in whodas_empties[1]:
                    if key in inner_dict and pd.isna(inner_dict[key]):
                        errors.append(f"In column '{dict_name}', in file '{file}', the field for '{key}' is empty (but others are filled)")
        
        # optional keys in CANS dictionary
        cans_empties = ['A_desc', 'B_desc', 'C_desc', 'D_desc']
        if dict_name == 'CANS':
            for key in cans_empties:
                if pd.isna(inner_dict[key]):
                    inner_dict[key] = ''
                        
        for key, item in inner_dict.items():
            if key not in whodas_empties[1] and key not in whodas_empties[0] and key not in cans_empties:
                if pd.isna(item):
                    errors.append(f"In column '{dict_name}', in file '{file}', the field for '{key}' is empty")                  
    
    return errors

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
            
            # Only keep indices that are not NaN and their corresponding values
            valid_indices = df[dict_name][df[dict_name].notna()]  # Indices without NaN
            valid_values = df[col][df[dict_name].notna()]
            
            master[dict_name] = pd.Series(valid_values.values, index=valid_indices).to_dict() # place dictionary in dictionary
            master[dict_name] = {k:int(v) if isinstance(v, float) and not math.isnan(v) else v for k, v in master[dict_name].items()} # convert values to integers where possible 
    
    master['GENERAL']['patient_name'] = master['GENERAL']['patient_first_name'] + " " + master['GENERAL']['patient_surname']
    master['GENERAL']['date'] = pd.to_datetime(master['GENERAL']['date']).strftime('%d/%m/%y')
    
    # calculate age
    today = datetime.now()
    age = today.year - master['GENERAL']['DOB'].year
    if (today.month, today.day) < (master['GENERAL']['DOB'].month, master['GENERAL']['DOB'].day):
        age -= 1
    master['GENERAL']['age'] = age
    
    return master # which contains all info needed to fill forms

def fill_textboxes(general_values:dict, form_values:dict, template):
    """
    Fills textbox values in pdf based on values in dictionary.
    """
    for page in template:
        for field in page.widgets(): # iterate through fields on each page
            key = field.field_name
            if field.field_type != fitz.PDF_WIDGET_TYPE_CHECKBOX:
                if key in form_values: # add form values
                    field.field_value = str(form_values[key])
                    field.update()
                if key in general_values: # add general values
                    field.field_value = str(general_values[key])
                    field.update()
                try: # in case form_values has float valued keys (for LSP)
                    if float(key) in form_values:
                        field.field_value = str(int(form_values[float(key)]))
                        field.update()
                except Exception:
                    pass
    return template

def highlight_box(x0, y0, x1, y1, page):
    """
    Highlights the area on template defined by coords x0, y0, x1, y1
    """
    rect = fitz.Rect(x0, y0, x1, y1) # create rectangle object
    highlight = page.add_highlight_annot(rect) # Add a highlight annotation to the defined box
    highlight.update() # save changes
    return page
    
def highlight_text(string:str, template, ins_no=0, case_sensitive=True):
    """
    Highlights the text on template defined by string
    """ 
    
    for page_num in range(template.page_count): # search each page
        page = template.load_page(page_num)
        
        text_instances = page.search_for(string) # find all instances of string
        
        if case_sensitive == True: # filter if case-sensitive
            final_instances = []
            for inst in text_instances:                
                if string in page.get_text("text", clip=inst): # compare 
                    final_instances.append(inst)
        else:
            final_instances = text_instances
        
        
        # add highlight and update for each string
        if ins_no < len(final_instances):
            highlight = page.add_highlight_annot(final_instances[ins_no])
            highlight.update()   
    
    return template     

def fill_WHODAS(general_values:dict, form_values:dict):
    """
    Inserts values from dictionary in correct fields in WHODAS pdf file.
    """
    template = fitz.open('forms/WHODAS.pdf')
            
    # calculate extra fields
    for i in range(1,7): 
        if i != 5:
            number_params = sum(1 for key, _ in form_values.items() if key.startswith('D' + str(i)))
            total = sum(value for key, value in form_values.items() if key.startswith('D' + str(i)))
            form_values[str(i) + '_overall'] = total
            form_values[str(i) + '_avg'] = round(total/number_params , 1)
            form_values[str(i) + '_percent'] = str(round((total/(number_params * 5)) * 100 , 1)) + "%"
    form_values['5_overall'] = form_values['D51'] + form_values['D52'] + form_values['D53'] + form_values['D51']
    form_values['5_avg'] = round(form_values['5_overall'] / 4, 1)
    form_values['5_percent'] = str(round((form_values['5_overall'] / 20) * 100, 1)) + '%'
    form_values['5_overall2'] = form_values['D55'] + form_values['D56'] + form_values['D57'] + form_values['D58']
    form_values['5_avg2'] = round(form_values['5_overall2'] / 4, 1)
    form_values['5_percent2'] = str(round((form_values['5_overall2'] / 20) * 100, 1)) + '%'
    
    # if section 2 of 5 is N/A and left empty
    if math.isnan(form_values['5_overall2']):
        form_values['D55'], form_values['D56'], form_values['D57'], form_values['D58'], form_values['5_avg2'], form_values['5_overall2'], form_values['5_percent2']  = "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"
        page = template.load_page(1)
        page.draw_line((26, 363), (583.7, 209.3), width=2)
        
        # find total values
        form_values['total'] = form_values['1_overall'] + form_values['2_overall'] + form_values['3_overall'] + form_values['4_overall'] + form_values['5_overall'] + form_values['6_overall'] + 20
    else:
        form_values['total'] = form_values['1_overall'] + form_values['2_overall'] + form_values['3_overall'] + form_values['4_overall'] + form_values['5_overall'] + form_values['5_overall2'] + form_values['6_overall']
        
    form_values['avg'] = round(form_values['total'] / 36, 1)
    form_values['percent'] = 'Total Score: ' + str(round((form_values['total'] / 180) * 100, 1)) + '%'
    
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

    # add in totals to dictionary
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

    # define coordinates for highlighting
    x_desc = [638.5,815.8] # for descriptions on right of page
    y_desc = [121.7,132.5,155.5,178.6,235.4,258.5,281.5,304.6,360.7]
    
    page = template.load_page(0) # higlights on page 1
    
    # calculate CANS level
    if form_values['A_subtotal'] < 4:
        if form_values['B_subtotal'] >= 4:
            form_values['total'] = 4.2
            page = highlight_box(x_desc[0], y_desc[3], x_desc[1], y_desc[4], page)
        elif form_values['C_subtotal'] >= 4:
            form_values['total'] = 4.1
            page = highlight_box(x_desc[0], y_desc[3], x_desc[1], y_desc[4], page)
        elif form_values['C_subtotal'] == 3:
            form_values['total'] = 3
            page = highlight_box(x_desc[0], y_desc[4], x_desc[1], y_desc[5], page)
        elif form_values['C_subtotal'] == 2:
            form_values['total'] = 2 
            page = highlight_box(x_desc[0], y_desc[5], x_desc[1], y_desc[6], page)
        elif form_values['C_subtotal'] == 1:
            form_values['total'] = 1
            page = highlight_box(x_desc[0], y_desc[6], x_desc[1], y_desc[7], page)
        else:
            form_values['total'] = 0  
            page = highlight_box(x_desc[0], y_desc[7], x_desc[1], y_desc[8], page)  
    elif form_values['A_subtotal'] == 4:
        form_values['total'] = 4.3
        page = highlight_box(x_desc[0], y_desc[3], x_desc[1], y_desc[4], page)
    elif form_values['A_subtotal'] == 5:
        form_values['total'] = 5
        page = highlight_box(x_desc[0], y_desc[2], x_desc[1], y_desc[3], page)
    elif form_values['A_subtotal'] == 6:
        form_values['total'] = 6
        page = highlight_box(x_desc[0], y_desc[1], x_desc[1], y_desc[2], page)
    else:
        form_values['total'] = 7
        page = highlight_box(x_desc[0], y_desc[0], x_desc[1], y_desc[1], page)
    
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
            page = highlight_box(x[0] + 10, y[i], x[1] - 10, y[i + 1], page) # bring in the highlight slightly due to formatting
        elif score == 1:
            page = highlight_box(x[1] + 10, y[i], x[2] - 10, y[i + 1], page) # bring in the highlight slightly due to formatting
        elif score == 2:
            page = highlight_box(x[2] + 10, y[i], x[3] - 10, y[i + 1], page) # bring in the highlight slightly due to formatting
        elif score == 3:
            page = highlight_box(x[3] + 10, y[i], x[4] - 10, y[i + 1], page) # bring in the highlight slightly due to formatting
    
    # perform scoring
    form_values['a_score'] = form_values[1] + form_values[2] + form_values[3] + form_values[8]
    form_values['b_score'] = form_values[4] + form_values[5] + form_values[6] + form_values[9] + form_values[16]
    form_values['c_score'] = form_values[10] + form_values[11] + form_values[12]
    form_values['d_score'] = form_values[7] + form_values[13] + form_values[14] + form_values[15]
    
    form_values['total'] = form_values['a_score'] + form_values['b_score'] + form_values['c_score'] + form_values['d_score']
    form_values['total_100'] = str(round(form_values['total'] * 2.0833, 2)) + "/100"
    
    form_values['a_score'] = str(form_values['a_score']) + "/12"
    form_values['b_score'] = str(form_values['b_score']) + "/15"
    form_values['c_score'] = str(form_values['c_score']) + "/9" 
    form_values['d_score'] = str(form_values['d_score']) + "/12"
    
    # fill form
    template = fill_textboxes(general_values, form_values, template) # fill out textboxes

    return template

def fill_LAWTON(general_values:dict, form_values:dict):
    """
    Inserts values from dictionary in correct fields in LAWTON pdf file.
    """
    template = fitz.open('forms/LAWTON.pdf') # read in template pdf
    
    with open('lawton.txt', 'r') as file:     
        sections = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
        i = 0 # counter
        for line in file:
            options = line.split('/')
            
            for opt_line in options[int(form_values[sections[i]]) - 1].split('*'):
                template = highlight_text(opt_line, template, case_sensitive=False) # highlight relevant number for each column
            
            i += 1 # increment
    
    # calculate left side total
    form_values['left_total'] = 0
    if form_values['A'] != 4:
        form_values['left_total'] += 1
    if form_values['B'] == 1:
        form_values['left_total'] += 1
    if form_values['C'] == 1:
        form_values['left_total'] += 1
    if form_values['D'] != 5:
        form_values['left_total'] += 1
        
    # calcualte right side total
    form_values['right_total'] = 0
    if form_values['E'] != 3:
        form_values['right_total'] += 1
    if form_values['F'] <= 3:
        form_values['right_total'] += 1
    if form_values['G'] == 1:
        form_values['right_total'] += 1
    if form_values['H'] != 3:
        form_values['right_total'] += 1 
           
    # calculate total 
    form_values['total'] = form_values['left_total'] + form_values['right_total']
    
    template = fill_textboxes(general_values, form_values, template)
    
    return template
        
def fill_BBS(general_values:dict, form_values:dict):
    """
    Inserts values from dictionary in correct fields in LAWTON pdf file.
    """
    template = fitz.open('forms/BBS.pdf') # read in template pdf

    total = 0
    
    for page in template: # gather fields
        for field in page.widgets():
            key = field.field_name
            # Check if it's a checkbox
            if field.field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
                category, value = key.split('_') # gain values for dictionary
                if form_values[int(category)] == float(value):
                    total += int(value)
                    field.field_value = True # mark correct checkboxes
                    field.update()
    
    new_dict = {}
    new_dict['total'] = total # new dictionary for efficiency, don't search through checkboxes
    template = fill_textboxes({}, new_dict, template) # no general values on BBS form
    
    return template

def fill_LEFS(general_values:dict, form_values:dict):
    """
    Inserts values from dictionary in correct fields in LEFS pdf file.
    """
    template = fitz.open('forms/LEFS.pdf') # read in template pdf
    
    x = [394,407,476,492,546,562,614,628,681,694]
    y = [199,211,223,235.5,249,261,274,286.5,299,312,324,337,349,362,374,386,399,412.5,425,438,449]

    # highlight correct box for each row
    page = template.load_page(0)
    pw = page.rect.width # page width
    total = 0 # track total value
    for i in range(20):
        score = form_values[i + 1]
        if score == 0:
            page = highlight_box(y[i] + 2, pw - x[1], y[i + 1] - 2, pw - x[0], page) # bring in the highlight slightly due to formatting
        elif score == 1:
            page = highlight_box(y[i] + 2, pw - x[3], y[i + 1] - 2, pw - x[2], page) # bring in the highlight slightly due to formatting
        elif score == 2:
            page = highlight_box(y[i] + 2, pw - x[5], y[i + 1] - 2, pw - x[4], page) # bring in the highlight slightly due to formatting
        elif score == 3:
            page = highlight_box(y[i] + 2, pw - x[7], y[i + 1] - 2, pw - x[6], page) # bring in the highlight slightly due to formatting
        elif score == 4:
            page = highlight_box(y[i] + 2, pw - x[9], y[i + 1] - 2, pw - x[8], page) # bring in the highlight slightly due to formatting
        total += score
    
    new_dict = {} # save in new to avoid passing form_values for efficiency
    new_dict['total'] = total
    
    template = fill_textboxes(general_values, new_dict, template) # fill other values
    
    return template

def fill_FRAT(general_values:dict, form_values:dict):
    """
    Inserts values from dictionary in correct fields in LEFS pdf file.
    """
    template = fitz.open('forms/FRAT.pdf') # read in template pdf

    # define coordinates for 'Part 1'
    x = [504, 521]
    y = [(203,213,224,236,248),(250,260.5,272,284,295),(297,307,318,330,341),(344,354,366,377,388)] # each tuple represents one question

    # highlight correct box for each row
    page = template.load_page(0) # load FRAT page
    
    total = 0 # track total
    
    key = 'Recent Falls' # key using 2,4,6,8 scale
    score = form_values[key]
    if score == 2:
        page = highlight_box(x[0], y[0][0], x[1], y[0][1], page) # choose correct y tuple and relevant values from tuple
    elif score == 4:
        page = highlight_box(x[0], y[0][1], x[1], y[0][2], page) 
    elif score == 6:
        page = highlight_box(x[0], y[0][2], x[1], y[0][3], page) 
    elif score == 8:
        page = highlight_box(x[0], y[0][3], x[1], y[0][4], page)    
    total += score
        
    keys = ['Medications', 'Psychological', 'Cognitive Status'] # keys using 1-4 scale
    for i in range(1, len(keys) + 1):
        score = form_values[keys[i - 1]]
        if score == 1:
            page = highlight_box(x[0], y[i][0], x[1], y[i][1], page) # choose correct y tuple and relevant values from tuple
        elif score == 2:
            page = highlight_box(x[0], y[i][1], x[1], y[i][2], page) 
        elif score == 3:
            page = highlight_box(x[0], y[i][2], x[1], y[i][3], page) 
        elif score == 4:
            page = highlight_box(x[0], y[i][3], x[1], y[i][4], page) 
        total += score
    
    for field in page.widgets(): # ineffiency to improve?
        key = field.field_name
        # Check if it's a checkbox
        if field.field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
            if form_values[key] == 'Y':
                field.field_value = True # mark checkbox
                field.update()
    
    form_values['total'] = total
    
    # highlight overall risk status
    x = [216.5,248,267,318,342,374]
    y = [502,514]
    
    # final fall status
    if form_values['auto_high_1'] == 'Y' or form_values['auto_high_2'] == 'Y' or total >= 16:
        page = highlight_box(x[4], y[0], x[5], y[1], page)
    elif total >= 5 and total <= 11:
        page = highlight_box(x[2], y[0], x[3], y[1], page)
    elif total >= 12 and total <= 15:
        page = highlight_box(x[0], y[0], x[1], y[1], page)

    # fill text fields
    template = fill_textboxes(general_values, form_values, template)
    
    return template
    
def fill_HONOS(general_values:dict, form_values:dict):
    """
    Inserts values from dictionary in correct fields in HONOS pdf file.
    """
    template = fitz.open('forms/HONOS.pdf') # read in template pdf  
    
    with open('honos.txt', 'r') as file: # read responses
        responses = file.readlines()
        
    total = 0 # track total score
    for i in range(len(responses)):
        options = responses[i].split('_') # different options
        
        value = form_values[i + 1] # value for current question
        opt = options[value]
        total += value # increment total 
        
        for line in opt.split('*'): # in case split over multiple lines
            # account for cases where line appears multiple times in the document
            if line == 'No problems of this kind during the period rated':
                instance = i
            else:
                instance = 0
            
            template = highlight_text(line, template, instance, case_sensitive=False) # highlight each line of value

    # specifications for question 8
    for letter in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
        if form_values[letter] == 'Y':
            if letter == 'A':
                template = highlight_text('A phobic', template)
                template = highlight_text('A,', template)
            elif letter == 'B':
                template = highlight_text('B anxiety', template)
                template = highlight_text('B,', template)
            elif letter == 'C':
                template = highlight_text('C obsessive-compulsive', template)
                template = highlight_text('C,', template)
            elif letter == 'D':
                template = highlight_text('D stress', template)
                template = highlight_text('D,', template)
            elif letter == 'E':
                template = highlight_text('E dissociative', template)
                template = highlight_text('E,', template)
            elif letter == 'F':
                template = highlight_text('F somatoform', template)
                template = highlight_text('F,', template)
            elif letter == 'G':
                template = highlight_text('G eating', template)
                template = highlight_text('G,', template)
            elif letter == 'H':
                template = highlight_text('H sleep', template)
                template = highlight_text('H,', template) 
            elif letter == 'I':
                template = highlight_text('I sexual', template)
                template = highlight_text('I,', template)
            elif letter == 'J':
                template = highlight_text('J other', template)
                template = highlight_text('J)', template)    
                              
    # to add in total
    form_values['total'] = str(total) + '/48'
    
    # fill in textboxes
    template = fill_textboxes(general_values, form_values, template)
    
    return template
    
def fill_WHODASKIDS(general_values:dict, form_values:dict):
    """
    Inserts values from dictionary in correct fields in HONOS pdf file.
    """
    
    template = fitz.open('forms/WHODASKIDS.pdf') # read in template pdf
    
    # calculate extra fields
    for i in range(1,7): 
        if i != 5:
            number_params = sum(1 for key, _ in form_values.items() if isinstance(key, float) and key // 10 == i) # number of keys in section
            total = sum(value for key, value in form_values.items() if isinstance(key, float) and key // 10 == i)
            form_values[str(i) + '_total'] = total
            form_values[str(i) + '_avg'] = round((total / (number_params * 5) * 100) , 1) # calculate percentage
            
    form_values['5_total'] = form_values[51] + form_values[52] + form_values[53] + form_values[54] # separate section 5 in case part 2 not included    
    form_values['5_total2'] = form_values[55] + form_values[56] + form_values[57] + form_values[58] + form_values[59]
    
    # if section 2 of 5 is N/A and left empty
    if math.isnan(form_values['5_total2']):
        form_values['55'], form_values['56'], form_values['57'], form_values['58'], form_values['59'], form_values['5_total2'], form_values['5_avg2'] = "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"
        page = template.load_page(1)
        page.draw_line((36.5,476.2), (505, 337), width=2)
        total = form_values['1_total'] + form_values['2_total'] + form_values['3_total'] + form_values['4_total'] + form_values['5_total'] + form_values['6_total'] + 25
    else:
        total = form_values['1_total'] + form_values['2_total'] + form_values['3_total'] + form_values['4_total'] + form_values['5_total'] + form_values['5_total2'] + form_values['6_total'] 
        form_values['5_avg2'] = round((form_values['5_total2'] / 25) * 100, 1)
    form_values['5_avg'] = round((form_values['5_total'] / 20) * 100, 1)
    
    # find total values    
    form_values['percentage'] = "Score: " + str(round(total / 34, 2)) + "/5 = " + str(round(total/1.7, 1)) + "%"
    
    form_values['total'] = "Total: " + str(total) + "/170"
    # add in total strings
    for i in range(1,7): 
        if i != 5:
            number_params = sum(1 for key, _ in form_values.items() if isinstance(key, float) and key // 10 == i) # number of keys in section
            form_values[str(i) + '_total'] = str(form_values[str(i) + '_total']) + "/" + str(number_params * 5)
            form_values[str(i) + '_avg'] = str(form_values[str(i) + '_avg']) + "%"
    
    # make strings for section 5
    form_values['5_total'] = str(form_values['5_total']) + "/20"
    form_values['5_avg'] = str(form_values['5_avg']) + "%"
    if form_values['5_total2'] != 'N/A':
        form_values['5_total2'] = str(form_values['5_total2']) + "/25"
        form_values['5_avg2'] = str(form_values['5_avg2']) + "%"
    
    # fill in textboxes
    template = fill_textboxes(general_values, form_values, template)
    return template

def fill_CASP(general_values:dict, form_values:dict):
    """
    Inserts values from dictionary in correct fields in CASP pdf file.
    """
    
    template = fitz.open('forms/CASP.pdf') # read in template pdf
    
    form_values['1_summary'] = 0
    form_values['2_summary'] = 0
    form_values['3_summary'] = 0
    form_values['4_summary'] = 0
    
    for page in template: # gather fields
        for field in page.widgets():
            key = field.field_name
            # Check if it's a checkbox
            if field.field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
                category, value = key.split('_') # gain values for dictionary
                if form_values[int(category)] == float(value):
                    if int(category) >= 1 and int(category) <= 6:
                        form_values['1_summary'] += int(value)
                    elif int(category) >= 6 and int(category) <= 10:
                        form_values['2_summary'] += int(value)
                    elif int(category) >= 10 and int(category) <= 15:
                        form_values['3_summary'] += int(value)
                    elif int(category) >= 15 and int(category) <= 20:
                        form_values['4_summary'] += int(value)
                    field.field_value = True # mark correct checkboxes
                    field.update()
    
    total = form_values['1_summary'] + form_values['2_summary'] + form_values['3_summary'] + form_values['4_summary'] 
    form_values['total'] = 'Total: ' + str(total) + '/80 = ' + str(round((total/80)*100, 2)) + '%'

    form_values['1_summary'] = str(form_values['1_summary']) + '/24'
    form_values['2_summary'] = str(form_values['2_summary']) + '/16'
    form_values['3_summary'] = str(form_values['3_summary']) + '/20'
    form_values['4_summary'] = str(form_values['4_summary']) + '/20'
    
    template = fill_textboxes(general_values, form_values, template)
    
    return template
    
def produce_output(master:dict[dict]):
    """
    Calls fill_form function for each dictionary in dictionaries and combines pdfs to final file. 
    """
    combined = fitz.open() # new document
    for key in master.keys():
        if key != 'GENERAL':
            try: # catch errors
                function_name = globals().get(f"fill_{key}")
                if function_name:
                    filled_form = function_name(master['GENERAL'], master[key])
                    rendered_pdf = render_to_image(filled_form) # this is a workaround to fuse field values to page 
                    combined.insert_pdf(rendered_pdf) # append to combined
            except Exception as e:
                raise Exception(f"Error in form {key}: {str(e)}")
    return combined

@app.route('/')
def index():
    """Render the upload page."""
    return render_template('upload.html')

@app.route('/download-template')
def download_template():
    """Serve the Excel template for download."""
    template_path = 'template.xlsx'  # Path to the Excel template
    return send_file(template_path, as_attachment=True)

@app.route('/upload', methods=['POST'])
def upload_files():
    """Handle file upload and return a zip file of processed PDFs."""
    # Check if the files were included in the request
    if 'files[]' not in request.files:
        return "No file part", 400

    files = request.files.getlist('files[]')

    # Validate file uploads
    if not files or all(f.filename == '' for f in files):
        return "No selected file", 400

    # Create an in-memory zip file
    memory_file = io.BytesIO()
    errors = []
    
    with zipfile.ZipFile(memory_file, 'w') as zf:
        for file in files:
            if file and file.filename.endswith('.xlsx'):
                # Ensure the filename is secure
                filename = secure_filename(file.filename)

                # Read the Excel file
                master = read_excel(file.stream)  # Function to read the Excel file
                
                # Validate the file contents
                errors.extend(validate_columns(master, file.filename))  # Updated validation that allows trailing empty rows
                
                if not errors: # prevent errors
                    final_document = produce_output(master)  # Function to generate the PDF from the DataFrame

                    # Store the PDF in memory
                    pdf_stream = io.BytesIO()
                    final_document.save(pdf_stream)
                    pdf_stream.seek(0)  # Reset the stream position

                    # Add the PDF to the zip file
                    pdf_filename = filename.replace('.xlsx', '')
                    zf.writestr(f'{pdf_filename}.pdf', pdf_stream.read())


    memory_file.seek(0)  # Reset the in-memory zip file position

    if errors:
        temp = errors
        errors = []
        return jsonify({"errors": temp}), 400 # flatten errors
    else:
        # Return the zip file as a downloadable response
        return send_file(memory_file, download_name='processed_files.zip', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)