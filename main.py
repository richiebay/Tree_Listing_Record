import PySimpleGUI as sg
import pandas as pd

# Defining headers for the table and the data frame.
headers = ['Officers', 'Species', 'DBH', 'Height Category', 'Diameter of Canopy', 'Health Classes',
           'Growth Environment', 'Longitude', 'Latitude']
numeric_fields = ['-dbh-', '-diameter-', '-longitude-', '-latitdude-']
# Initializing Data frame.
df = pd.DataFrame(columns=headers)

# Calculating column width based on the length of each header's string.
col_widths = list(map(lambda x: len(x) + 3, headers))

# Defining each entry as a list of text field and it's appropriate input.
# Giving keys is essential to retrieve the values of elements inside the event loop.
officers = [sg.Text('Officers'),
            sg.Combo(['Vincenzo Rich', 'Callen Duggan', 'Damien Sheppard', 'Cherry Morales', 'Jackson Blake'],
                     key='-officers-')]
species = [sg.Text('Species'),
           sg.Combo(['black locust', 'black pine', 'common juniper', 'common yew', 'eurpean larch'], key='-species-'
                    )]
dbh = [sg.Text('DBH'), sg.InputText('0.0', size=(19, 20), justification='right', key='-dbh-')]
height = [sg.Text('Height Category'),
          sg.Combo(['<20ft', '20-40ft', '>40ft', '>50ft'], key='-height-')]
diameter = [sg.Text('Diameter of Canopy'), sg.InputText('0.0', size=(20, 20), justification='right', key='-diameter-')]
health = [sg.Text('Health Classes'),
          sg.Combo(['Excellent', 'Very Good', 'Good', 'Fair', 'Poor', 'Dead'], key='-health-'
                   )]
growth = [sg.Text('Growth Environment'),
          sg.Combo(['Good', 'Fair', 'Poor'], size=(6, 20), key='-growth-')]
longitude = [sg.Text('Longitude'), sg.InputText('0.0', size=(20, 20), justification='right', key='-longitude-')]
latitdude = [sg.Text('Latitude'), sg.InputText('0.0', size=(10, 20), justification='right', key='-latitdude-')]

table = [sg.Table(values=df.values.tolist(),
                  key='-LIST-',
                  headings=headers,
                  display_row_numbers=True,
                  auto_size_columns=False,
                  max_col_width=40,
                  def_col_width=14,
                  justification='middle',
                  row_height=35,
                  col_widths=col_widths, enable_click_events=True,
                  num_rows=min(25, len(df.values.tolist()) + 5))]
# Designing the gui , in PySimple GUI , the gui is divided to rows and columns.
row1c2 = [sg.Text('Search by'),
          sg.Combo(values=headers, key="-searchbycombo-", default_value='Select an option', size=(14, 20)),
          sg.InputText(size=(20, 20), justification='right', key='-searchbytext-'),
          sg.Button('Search', size=(10, 1))]

row2c2 = [sg.Text('        '), sg.Button('Export', size=(20, 2)), sg.Button('Show All', size=(20, 2))]

row1 = officers + species
row2 = growth + height
row3 = health + dbh
row4 = diameter
row5 = longitude + latitdude
row6 = [sg.Button('New Entry', size=(15, 2)), sg.Button('Update', size=(15, 2)), sg.Button('Delete', size=(15, 2))]
row7 = table

row8 = [sg.Button('Close')]
col1 = [row1, row2, row3, row4, row5, row6]
col2 = [row1c2, row2c2]

layout = [[sg.Text('        '), sg.Column(col1), sg.Text('        '), sg.VerticalSeparator(), sg.Text('        '),
           sg.Col(col2)],
          row7,
          row8]

# Initializing the window object with the recently created rows and columns.
window = sg.Window('Tree Listing Record', layout, grab_anywhere=False)


# This is a search function that searches for certain value in a data frame given its key.
def search_function(df, column_name, search_value):
    values_to_search_in = list(df[column_name].values)
    print(values_to_search_in)
    found = []
    for index, value in enumerate(values_to_search_in):
        if search_value == value:
            found.append(df.values.tolist()[index])
        else:
            continue
    return found


# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Close':  # if user closes window or clicks cancel
        break
    elif event == 'New Entry':
        try:
            for field in numeric_fields:
                try:
                    float(values[field])
                except:
                    sg.popup_error("Value of " + field + "must be number",title="Error",auto_close=True)
                    raise Exception("Invalid entry")
            for i in range(5):
                if list(values.values())[i] == "":
                    sg.popup_error("Values cannot be empty",title="Error",auto_close=True)
                    raise Exception("Invalid entry")
            current_window_values = [values['-officers-'], values['-species-'], values['-dbh-'], values['-height-'],
                                     values['-diameter-'],
                                     values['-health-'], values['-growth-'], values['-longitude-'],
                                     values['-latitdude-']]
            temp_df = pd.DataFrame([current_window_values], columns=headers)
            df = pd.concat([df, temp_df], ignore_index=True)
            window['-LIST-'].update(df.values.tolist())

            # Returning values to default values
            officers[1].update(value='')
            species[1].update(value='')
            growth[1].update(value='')
            height[1].update(value='')
            health[1].update(value='')
            dbh[1].update(value='0.0')
            diameter[1].update(value='0.0')
            longitude[1].update(value='0.0')
            latitdude[1].update(value='0.0')
        except:
            continue

    #     print(values['-LIST-'][0])
    elif event == 'Export':
        if not (df.values.tolist()):
            sg.popup_error("No values found to export.",title="Error",auto_close=True)
            continue
        export = [
            [sg.InputText(key='file_save_as_input', enable_events=True),
             sg.FileSaveAs(key='file_save_as_key', initial_folder='/tmp')
             ]]
        export_event, export_values = sg.Window('POPUP', export, modal=True).read(close=True)
        try:
            foldername = export_values['file_save_as_input']
            if '.xlsx' in foldername:
                df.to_excel(foldername)
            else:
                df.to_excel(foldername + '.xlsx')
        except:
            continue

    elif event == 'Search':
        try:
            column_name = values['-searchbycombo-']
            if column_name not in headers:
                continue
            value_to_look_for = values['-searchbytext-']
            found_values = search_function(df, column_name, value_to_look_for)
            if len(found_values) > 0:
                window['-LIST-'].update(found_values)
            else:
                sg.popup_error('Entry Not Found',title="Error",auto_close=True)
        except:
            continue

    elif event == 'Show All':
        window['-LIST-'].update(df.values.tolist())

    elif event == 'Update':
        try:
            row = values['-LIST-'][0]
            current_window_values = [values['-officers-'], values['-species-'], values['-dbh-'], values['-height-'],
                                     values['-diameter-'],
                                     values['-health-'], values['-growth-'], values['-longitude-'],
                                     values['-latitdude-']]
            df.iloc[row] = current_window_values
            window['-LIST-'].update(df.values.tolist())
            sg.popup_notify("Table Entry Updated With Values In The New Entry.", display_duration_in_ms=200,
                            fade_in_duration=1000,title="Success",auto_close=True)
        except:
            sg.popup_error("Please select an entry to update.",title="Error",auto_close=True)

    elif event == 'Delete':
        try:
            row = values['-LIST-'][0]
            df = df.drop(row)
            df.reset_index(drop=True, inplace=True)
            window['-LIST-'].update(df.values.tolist())
            sg.popup_notify("Table Entry Deleted.", display_duration_in_ms=100, fade_in_duration=1000, title="Success")
        except:
            sg.popup_error("Please Select An Entry To Delete.", title="Error",auto_close=True)

    elif len(event) > 1:
        try:
            if event[0] == '-LIST-':
                row = event[2][0]
                values_to_update = df.values.tolist()[row]
                # values['-officers-'] = 'certain_value'
                officers[1].update(value=values_to_update[0])
                species[1].update(value=values_to_update[1])
                growth[1].update(value=values_to_update[6])
                height[1].update(value=values_to_update[3])
                health[1].update(value=values_to_update[5])
                dbh[1].update(value=values_to_update[2])
                diameter[1].update(value=values_to_update[4])
                longitude[1].update(value=values_to_update[7])
                latitdude[1].update(value=values_to_update[8])
        except:
            continue

window.close()
