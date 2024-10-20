import io
from datetime import datetime
import math
import zipfile

import pandas as pd
import fitz
from flask import Flask, request, send_file, render_template, jsonify, send_from_directory
from werkzeug.utils import secure_filename

app = Flask(__name__)

def validate_columns(master, file):
    """
    This function validates the master dictionary previously created and generates and returns a list of errors. 
    Validate that if any row in each dictionary (for each form) is filled, all subsequent rows must also be filled. 
    """
    
    error_messages = [] # to store all error messages
    
    # WHODAS files have a series of optional entries. They can either all be empty or all be full. This is validated here.
    whodas_empties = {}
    whodas_empties['WHODAS'] = ['D55', 'D56', 'D57', 'D58']
    whodas_empties['WHODASKIDS'] = [55, 56, 57, 58, 59]

    # perform check for each dictionary
    for dict_name, inner_dict in master.items():
        
        # special check for optional WHODAS columns
        if dict_name == 'WHODAS' or dict_name == 'WHODASKIDS':     
            has_filled = any(not pd.isna(inner_dict[key]) for key in whodas_empties[dict_name] if key in inner_dict) # filled optional entries
            has_empty = any(pd.isna(inner_dict[key]) for key in whodas_empties[dict_name] if key in inner_dict) # empty optional entries
            
            # If there are filled and empty entries in the optional rows, return error
            if has_filled and has_empty:
                for key in whodas_empties[dict_name]: 
                    if key in inner_dict and pd.isna(inner_dict[key]): # attain empty key names
                        error_messages.append(f"In column '{dict_name}', the field for '{key}' is empty")
        
        # other dictionaries have optional columns without the need for further logic
        optional = {}
        optional['CANS'] = ['A_desc', 'B_desc', 'C_desc', 'D_desc']
        optional['HONOS'] = ['comment8']
        optional['FRAT'] = ['Other_desc']
        if dict_name in optional.keys():
            for key in optional[dict_name]:
                if pd.isna(inner_dict[key]):
                    inner_dict[key] = '' # empty string assigned to prevent NaN
        
        # final check for NaN rows
        for key, item in inner_dict.items():
            if key not in whodas_empties['WHODAS'] and key not in whodas_empties['WHODASKIDS']: # checks already performed for optional WHODAS columns
                if pd.isna(item):
                    error_messages.append(f"In column '{dict_name}', the field for '{key}' is empty")                 
    
    return error_messages

def render_to_image(filled_form):
    """
    Renders the filled PDF form to images and saves them as new PDFs.
    """
    
    temp_pdf = fitz.open()  # Create a new PDF 
    
    # copy each page to new pdf in image form
    for page_number in range(len(filled_form)):
        
        page = filled_form[page_number]
        
        # Render page to an image
        pix = page.get_pixmap(matrix=fitz.Matrix(3,3))  # Zoom for better quality
        
        img_pdf = fitz.open()  # New PDF for this page
        img_page = img_pdf.new_page(width=pix.width, height=pix.height)  # Create a new page
        img_page.insert_image(img_page.rect, stream=pix.tobytes())  # Insert image into the new page
        
        temp_pdf.insert_pdf(img_pdf)  # Insert the image PDF into the temp PDF

    return temp_pdf

def read_excel(excel):
    """
    Reads in path to excel file and populates relevant dictionaries with values.
    Stores dictionaries in master dictionary and returns master dictionary. 
    """

    df = pd.read_excel(excel) # read excel into pandas df
    master = {} # initialise master dictionary
    
    for col in df.columns:
        if (df[col].notna().any() and 'values' in col.lower()) or col == 'GENERAL Values': # find dictionaries to populate    
            dict_name = col.replace(' Values', '') # dictionary key name
            
            # Only keep indices that are not NaN and their corresponding values
            valid_indices = df[dict_name][df[dict_name].notna()]  # Indices without NaN
            valid_values = df[col][df[dict_name].notna()]
            
            master[dict_name] = pd.Series(valid_values.values, index=valid_indices).to_dict() # place dictionary in dictionary
            master[dict_name] = {k:int(v) if isinstance(v, float) and not math.isnan(v) else v for k, v in master[dict_name].items()} # convert values to integers where possible 
            master[dict_name] = {int(k) if isinstance(k, float) and not math.isnan(k) else k: int(v) if isinstance(v, float) and not math.isnan(v) else v for k, v in master[dict_name].items()} # convert key to integer if float
    temp = master['GENERAL'].copy() # to iterate over so master['GENERAL'] can change
        
    for key, item in temp.items(): # all GENERAL columns to be accounted for, replaced with empty strings if no values entered.
        if pd.isna(item):
            master['GENERAL'][key] = '' # empty string for NaN 
        else: # if not empty
            
            if key == 'date': # convert date to DD/MM/YY format
                master['GENERAL']['date'] = pd.to_datetime(master['GENERAL']['date']).strftime('%d/%m/%y')
            
            if key == 'DOB': # calculate age
                today = datetime.today()
                age = today.year - master['GENERAL']['DOB'].year
                
                # subtract a year if birthday has not occured this year
                if (today.month, today.day) < (master['GENERAL']['DOB'].month, master['GENERAL']['DOB'].day):
                    age -= 1
                
                master['GENERAL']['age'] = age # assign age to dictionary
                master['GENERAL']['DOB'] = pd.to_datetime(master['GENERAL']['DOB']).strftime('%d/%m/%y') # format DOB
        
    # combine first and last name for full patient_name
    master['GENERAL']['patient_name'] = master['GENERAL']['patient_first_name'] + " " + master['GENERAL']['patient_surname']

    return master # contains all info needed to fill forms

def fill_textboxes(general_values:dict, form_values:dict, template):
    """
    Fills textbox values in pdf based on values in dictionaries general_values and form_values.
    """
    
    for page in template:
        for field in page.widgets(): # iterate through fields on each page
            
            if field.field_type != fitz.PDF_WIDGET_TYPE_CHECKBOX: # if field is not checkbox
                
                key = field.field_name # field name
                if key in form_values: # add form values to template
                    field.field_value = str(form_values[key])
                if key in general_values: # add general values to template
                    field.field_value = str(general_values[key])
                field.update()
                
                try: # for integer type form_value values
                    if int(key) in form_values:
                        field.field_value = str(form_values[int(key)])
                        field.update()
                except Exception:
                    continue
                
    
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
    Highlights the text on template defined by string. By default, case senstive search and can optionally add a number for which instance to highlight. 
    """ 
        
    for page_num in range(template.page_count): # search each page
        page = template.load_page(page_num) # load page
        
        text_instances = page.search_for(string) # find all instances of string on page
    
    
        if case_sensitive == True: # filter if case-sensitive
                
            final_instances = [] # stores filtered instances
                
            for inst in text_instances:                
                if string in page.get_text("text", clip=inst): # compare string to instances
                    final_instances.append(inst)
        else:
            final_instances = text_instances # not case-sensitive search by default with fitz
        
        
        # add highlight and update for each string
        if ins_no < len(final_instances):
            highlight = page.add_highlight_annot(final_instances[ins_no])
            highlight.update()
            return template
        else:
            ins_no -= len(final_instances)
    
    return template     

def fill_WHODAS(general_values:dict, form_values:dict):
    """
    Inserts values from dictionary in correct fields in WHODAS pdf file.
    """
    
    template = fitz.open('forms/WHODAS.pdf')
            
    # calculate extra fields
    for i in range(1,7): 
        if i != 5: # 5 is an edge case
            
            number_params = sum(1 for key, _ in form_values.items() if key.startswith('D' + str(i)))
            total = sum(value for key, value in form_values.items() if key.startswith('D' + str(i)))
            
            form_values[str(i) + '_overall'] = total
            form_values[str(i) + '_avg'] = round(total/number_params , 1)
            form_values[str(i) + '_percent'] = str(round((total/(number_params * 5)) * 100 , 1)) + "%"
    
    # calculate values for section 5 part 1        
    form_values['5_overall'] = form_values['D51'] + form_values['D52'] + form_values['D53'] + form_values['D54']
    form_values['5_avg'] = round(form_values['5_overall'] / 4, 1)
    form_values['5_percent'] = str(round((form_values['5_overall'] / 20) * 100, 1)) + '%'
    
    # calculate values for section 5 part 2
    form_values['5_overall2'] = form_values['D55'] + form_values['D56'] + form_values['D57'] + form_values['D58']
    form_values['5_avg2'] = round(form_values['5_overall2'] / 4, 1)
    form_values['5_percent2'] = str(round((form_values['5_overall2'] / 20) * 100, 1)) + '%'
    
    # if part 2 of 5 is N/A and left empty
    if math.isnan(form_values['5_overall2']):
        form_values['D55'], form_values['D56'], form_values['D57'], form_values['D58'], form_values['5_avg2'], form_values['5_overall2'], form_values['5_percent2']  = "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "" # will appear on document as N/A
        
        page = template.load_page(1)
        page.draw_line((26, 363), (583.7, 209.3), width=2) # cross out section
        
        # find total values, counting an extra 20 for part 2
        form_values['total'] = form_values['1_overall'] + form_values['2_overall'] + form_values['3_overall'] + form_values['4_overall'] + form_values['5_overall'] + form_values['6_overall'] + 20
    
    else: # total values is equal to all sections combined
        form_values['total'] = form_values['1_overall'] + form_values['2_overall'] + form_values['3_overall'] + form_values['4_overall'] + form_values['5_overall'] + form_values['5_overall2'] + form_values['6_overall']
    
    # calculate final extra values using total
    form_values['avg'] = round(form_values['total'] / 36, 1)
    form_values['percent'] = 'Total Score: ' + str(round((form_values['total'] / 180) * 100, 1)) + '%'
    
    # Fill in textboxes with values generated
    template = fill_textboxes(general_values, form_values, template)

    # Fill in checkboxes for male and female
    page = template.load_page(0)
    check = False # track if male or female selected
    for field in page.widgets():
        
        # Check if it's a checkbox
        if field.field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
            
            key = field.field_name
            
            # check if male
            if key == 'male' and general_values['gender'].lower() == 'm':
                field.field_value = True  # Set checkbox to checked
                check = True 
            
            # check if female
            elif key == 'female' and general_values['gender'].lower() == 'f':
                field.field_value = True  # Set checkbox to checked
                check = True 
            
            field.update()
            
            if check: # only one checkbox, for efficiency can skip rest of for loop once found
                break
    return template

def fill_WHODASKIDS(general_values:dict, form_values:dict):
    """
    Inserts values from dictionary in correct fields in HONOS pdf file.
    """
    
    template = fitz.open('forms/WHODASKIDS.pdf') # read in template pdf
    
    # calculate extra fields
    for i in range(1,7): 
        if i != 5: # section 5 is edge case
            number_params = sum(1 for key, _ in form_values.items() if isinstance(key, int) and key // 10 == i) # number of keys in section
            total = sum(value for key, value in form_values.items() if isinstance(key, int) and key // 10 == i)
            
            form_values[str(i) + '_total'] = total
            form_values[str(i) + '_avg'] = round((total / (number_params * 5) * 100) , 1) # calculate percentage

    
    # calcualte for section 5
    form_values['5_total'] = form_values[51] + form_values[52] + form_values[53] + form_values[54]
    form_values['5_total2'] = form_values[55] + form_values[56] + form_values[57] + form_values[58] + form_values[59]
    
    # if section 2 of 5 is N/A and left empty
    if math.isnan(form_values['5_total2']):
        form_values[55], form_values[56], form_values[57], form_values[58], form_values[59], form_values['5_total2'], form_values['5_avg2'] = "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"
        
        # cross out empty section
        page = template.load_page(1)
        page.draw_line((36.5,476.2), (505, 337), width=2)
        
        # calculate document total without part 2 of section 5
        total = form_values['1_total'] + form_values['2_total'] + form_values['3_total'] + form_values['4_total'] + form_values['5_total'] + form_values['6_total'] + 25
    else:
        # calculate document total with part 2 of section 5
        total = form_values['1_total'] + form_values['2_total'] + form_values['3_total'] + form_values['4_total'] + form_values['5_total'] + form_values['5_total2'] + form_values['6_total'] 
        form_values['5_avg2'] = round((form_values['5_total2'] / 25) * 100, 1)
    
    form_values['5_avg'] = round((form_values['5_total'] / 20) * 100, 1)
    
    # find total values    
    form_values['percentage'] = "Score: " + str(round(total / 34, 2)) + "/5 = " + str(round(total/1.7, 1)) + "%"
    form_values['total'] = "Total: " + str(total) + "/170"
    
    # add in strings for presentation on document
    for i in range(1,7): 
        if i != 5:
            number_params = sum(1 for key, _ in form_values.items() if isinstance(key, int) and key // 10 == i) # number of keys in section
            form_values[str(i) + '_total'] = str(form_values[str(i) + '_total']) + "/" + str(number_params * 5)
            form_values[str(i) + '_avg'] = str(form_values[str(i) + '_avg']) + "%"
    
    # make strings for section 5
    form_values['5_total'] = str(form_values['5_total']) + "/20"
    form_values['5_avg'] = str(form_values['5_avg']) + "%"
    if form_values['5_total2'] != 'N/A':
        form_values['5_total2'] = str(form_values['5_total2']) + "/25"
        form_values['5_avg2'] = str(form_values['5_avg2']) + "%"
        
    template = fill_textboxes(general_values, form_values, template)
    
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

    # tick the relevant checkboxes on the page and add to totals
    for page in template: # gather fields
        for field in page.widgets():
            # Check if it's a checkbox
            if field.field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
                
                key = field.field_name # field name
                question_no = int(key[1:]) # all checkboxes can be converted to integer
                
                if form_values[question_no].upper() == 'Y' and key[:1] == 'Y':
                    field.field_value = True  # Set checkbox to checked
                    
                    # add to totals
                    if question_no > 0 and question_no < 11:
                        form_values['A_subtotal'] += 1
                    elif question_no > 10 and question_no < 15:
                        form_values['B_subtotal'] += 1
                    elif question_no > 14 and question_no < 26:
                        form_values['C_subtotal'] += 1
                    elif question_no > 25 and question_no < 29:
                        form_values['D_subtotal'] += 1                    
                          
                elif form_values[question_no].upper() == 'N' and key[:1] == 'N':
                    field.field_value = True  # Set checkbox to checked
                field.update()
    
    # calculate total 
    form_values['subtotal'] = form_values['A_subtotal'] + form_values['B_subtotal'] + form_values['C_subtotal'] + form_values['D_subtotal']

    # define coordinates for highlighting descriptions on right side of page
    x_desc = [638.5,815.8] 
    y_desc = [121.7,132.5,155.5,178.6,235.4,258.5,281.5,304.6,360.7]
    
    page = template.load_page(0) # higlights on page 1
    
    # calculate CANS level
    if form_values['A_subtotal'] < 4:
        if form_values['B_subtotal'] >= 4:
            form_values['total'] = 4.2
            page = highlight_box(x_desc[0], y_desc[3], x_desc[1], y_desc[4], page)
        elif form_values['C_subtotal'] >= 4: # check C subtotal
            form_values['total'] = 4.1
            page = highlight_box(x_desc[0], y_desc[3], x_desc[1], y_desc[4], page)
        elif form_values['C_subtotal'] == 3 or form_values['D_subtotal'] == 3:
            form_values['total'] = 3
            page = highlight_box(x_desc[0], y_desc[4], x_desc[1], y_desc[5], page)
        elif form_values['C_subtotal'] == 2 or form_values['D_subtotal'] == 2:
            form_values['total'] = 2 
            page = highlight_box(x_desc[0], y_desc[5], x_desc[1], y_desc[6], page)
        elif form_values['C_subtotal'] == 1 or form_values['D_subtotal'] == 1:
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
        
        # highlight scores using coordinate
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
    
    # total values and percentage
    form_values['total'] = form_values['a_score'] + form_values['b_score'] + form_values['c_score'] + form_values['d_score']
    form_values['total_100'] = str(round(form_values['total'] * 2.0833, 2)) + "/100"
    
    # turn into strings for presentation
    form_values['a_score'] = str(form_values['a_score']) + "/12"
    form_values['b_score'] = str(form_values['b_score']) + "/15"
    form_values['c_score'] = str(form_values['c_score']) + "/9" 
    form_values['d_score'] = str(form_values['d_score']) + "/12"
    
    template = fill_textboxes(general_values, form_values, template) # fill out textboxes

    return template

def fill_LAWTON(general_values:dict, form_values:dict):
    """
    Inserts values from dictionary in correct fields in LAWTON pdf file.
    """
    
    template = fitz.open('forms/LAWTON.pdf') # read in template pdf
    
    # use text document to highlight relevant rows for each question
    with open('forms/lawton.txt', 'r') as file:     
        sections = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
        i = 0 # counter
        for line in file:
            options = line.split('/') # each option separated by /
            
            # highlight each line, separated by *
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
    
    # fill in fields
    template = fill_textboxes(general_values, form_values, template)
    
    return template
        
def fill_BBS(_, form_values:dict):
    """
    Inserts values from dictionary in correct fields in LAWTON pdf file.
    """
    
    template = fitz.open('forms/BBS.pdf') # read in template pdf

    total = 0 # to increment for patient total score
    
    for page in template: # gather fields
        for field in page.widgets():
            
            # field name
            key = field.field_name
            
            if field.field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX: # Check if it's a checkbox
                
                category, value = key.split('_') # gain values for dictionary
                
                if form_values[int(category)] == float(value):
                    total += int(value) # increase total score
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
    
    # coordinates of boxes for highlighting
    x = [394,407,476,492,546,562,614,628,681,694]
    y = [199,211,223,235.5,249,261,274,286.5,299,312,324,337,349,362,374,386,399,412.5,425,438,449]

    page = template.load_page(0) # load specific page
    pw = page.rect.width # page width
    
    # track totals
    new_dict = {} # save in new dict to avoid passing all form_values for efficiency
    new_dict['total'] = 0
    new_dict['0_total'] = 0
    new_dict['1_total'] = 0
    new_dict['2_total'] = 0
    new_dict['3_total'] = 0
    new_dict['4_total'] = 0
    
    # highlight correct box for each score and increment totals
    for i in range(20):
        
        # find score
        score = form_values[i + 1] 
        
        if score == 0:
            page = highlight_box(y[i] + 2, pw - x[1], y[i + 1] - 2, pw - x[0], page) # bring in the highlight slightly due to formatting
        elif score == 1:
            page = highlight_box(y[i] + 2, pw - x[3], y[i + 1] - 2, pw - x[2], page) # bring in the highlight slightly due to formatting
            new_dict['1_total'] += 1
        elif score == 2:
            page = highlight_box(y[i] + 2, pw - x[5], y[i + 1] - 2, pw - x[4], page) # bring in the highlight slightly due to formatting
            new_dict['2_total'] += 2
        elif score == 3:
            page = highlight_box(y[i] + 2, pw - x[7], y[i + 1] - 2, pw - x[6], page) # bring in the highlight slightly due to formatting
            new_dict['3_total'] += 3
        elif score == 4:
            page = highlight_box(y[i] + 2, pw - x[9], y[i + 1] - 2, pw - x[8], page) # bring in the highlight slightly due to formatting
            new_dict['4_total'] += 4 
        new_dict['total'] += score # increment total score

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

    
    page = template.load_page(0) # load FRAT page
    total = 0 # track total
    
    # highlight correct box for each row
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
    total += score # increment score
        
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
        total += score # increment score
    
    # fill checkboxes
    for field in page.widgets():
       
        key = field.field_name # name of field
        
        # Check if it's a checkbox
        if field.field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
            if form_values[key] == 'Y':
                field.field_value = True # mark checkbox
                field.update()
    
    form_values['total'] = total
    
    # highlight overall risk status
    x = [216.5,248,267,318,342,374]
    y = [502,514]
    
    # calculate final fall status
    if form_values['auto_high_1'] == 'Y' or form_values['auto_high_2'] == 'Y' or total >= 16: # high falls risk
        page = highlight_box(x[4], y[0], x[5], y[1], page)
    
    elif total >= 5 and total <= 11: # medium falls risk
        page = highlight_box(x[2], y[0], x[3], y[1], page)
        
    elif total >= 12 and total <= 15: # low falls risk
        page = highlight_box(x[0], y[0], x[1], y[1], page)

    # fill text fields
    template = fill_textboxes(general_values, form_values, template)
    
    return template
    
def fill_HONOS(general_values:dict, form_values:dict):
    """
    Inserts values from dictionary in correct fields in HONOS pdf file.
    """
    template = fitz.open('forms/HONOS.pdf') # read in template pdf  
    
    with open('forms/honos.txt', 'r') as file: # read responses
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
            elif i + 1 == 9 and value == 1: # question 9 has a repeated option from previous question
                instance = 1
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

def fill_CASP(general_values:dict, form_values:dict):
    """
    Inserts values from dictionary in correct fields in CASP pdf file.
    """
    
    template = fitz.open('forms/CASP.pdf') # read in template pdf
    
    # total values
    form_values['1_summary'] = 0
    form_values['2_summary'] = 0
    form_values['3_summary'] = 0
    form_values['4_summary'] = 0
    
    for page in template: # gather fields
        for field in page.widgets():
            
            key = field.field_name
            
            
            if field.field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:# Check if it's a checkbox
                
                category, value = key.split('_') # gain values for dictionary
                
                # calculate totals
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
    
    # form totals
    total = form_values['1_summary'] + form_values['2_summary'] + form_values['3_summary'] + form_values['4_summary'] 
    form_values['total'] = 'Total: ' + str(total) + '/80 = ' + str(round((total/80)*100, 2)) + '%'

    # totals as full strings
    form_values['1_summary'] = str(form_values['1_summary']) + '/24'
    form_values['2_summary'] = str(form_values['2_summary']) + '/16'
    form_values['3_summary'] = str(form_values['3_summary']) + '/20'
    form_values['4_summary'] = str(form_values['4_summary']) + '/20'
    
    template = fill_textboxes(general_values, form_values, template)
    
    return template
    
def produce_output(master:dict[dict]):
    """
    Calls form filling function for each dictionary read in from excel and combines pdfs to final file. 
    """
    
    combined = fitz.open() # new document to return
    
    for key in master.keys():
        if key != 'GENERAL':
            function_name = globals().get(f"fill_{key}") # function to call to fill out form
            
            if function_name: # check function exists to prevent errors
                
                filled_form = function_name(master['GENERAL'], master[key])
                rendered_pdf = render_to_image(filled_form) # this is a workaround to fuse field values to page 
                combined.insert_pdf(rendered_pdf) # append to combined
                
    return combined

@app.route('/')
def index():
    """Render the upload page."""
    return render_template('upload.html')

@app.route('/download-template')
def download_template():
    """Serve the Excel template for download."""
    
    template_path = '../template.xlsx'  # Path to the Excel template
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
    errors = {}
    
    with zipfile.ZipFile(memory_file, 'w') as zf:
        for file in files:
            if file and file.filename.endswith('.xlsx'):
                # Ensure the filename is secure
                filename = secure_filename(file.filename)

                # Read the Excel file
                master = read_excel(file.stream)  # Function to read the Excel file
                
                # Validate the file contents
                error_list = validate_columns(master, file.filename)
                if error_list:
                    errors[file.filename] = error_list  # Updated validation that allows trailing empty rows
                
                try: # use try in case validation misses an error
                    if all(not lst for lst in errors.values()): # prevent errors
                        final_document = produce_output(master)  # Function to generate the PDF from the DataFrame

                        # Store the PDF in memory
                        pdf_stream = io.BytesIO()
                        final_document.save(pdf_stream)
                        pdf_stream.seek(0)  # Reset the stream position

                        # Add the PDF to the zip file
                        pdf_filename = filename.replace('.xlsx', '')
                        zf.writestr(f'{pdf_filename}.pdf', pdf_stream.read())
                except Exception:
                    errors = [f"There is an issue with {file.filename}. Please ensure the correct template has been used. If errors reoccur, redownload the template and try again."]


    memory_file.seek(0)  # Reset the in-memory zip file position
    
    if errors:
        return jsonify({"errors": errors}), 400
    else:
        # Return the zip file as a downloadable response
        return send_file(memory_file, download_name='processed_files.zip', as_attachment=True)
    
@app.route('/download-form/<form_name>')
def download_form(form_name):
    """
    Route to download specific forms.
    """
    # Map form_name to actual file paths
    form_files = {
        'whodas': 'WHODAS.pdf',
        'whodas-youth': 'WHODASKIDS.pdf',
        'cans': 'CANS.pdf',
        'lsp': 'LSP.pdf',
        'lawton-brody-iadl': 'LAWTON.pdf',
        'lefs': 'LEFS.pdf',
        'berg-balance-scale': 'BBS.pdf',
        'frat': 'FRAT.pdf',
        'honos': 'HONOS.pdf',
        'casp': 'CASP.pdf'
    }
    
    # Ensure the form exists
    if form_name in form_files:
        return send_from_directory(directory='forms', path=form_files[form_name], as_attachment=True)
    else:
        return "Form not found", 404

if __name__ == '__main__':
    app.run(debug=True)