import pandas as pd
import fitz

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

def produce_output(dictionaries:dict[dict]):
    """
    Calls fill_form function for each dictionary in dictionaries and combines pdfs to final file. 
    """
    pass

master = read_excel('Book1.xlsx')

#for key in master.keys():
#    if key != 'GENERAL':
#        writer = PdfWriter()
#        writer.addpages(fill_form(master['GENERAL'], master[key], 'forms/'+ key + '.pdf'))
#        writer.write(key + '.pdf')

output_template = fill_CANS(master['GENERAL'], master['CANS'])
output_template.save('CANS.pdf')