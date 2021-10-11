import os
import platform
import pandas as pd
import PySimpleGUI as sg
# Logo need to be implemented
#from logo.logo import get_logo

# Find the platform and setup environment
# Not yet implemented
if platform.system() == "Darwin":
    # Do Mac Setup
    pass
elif platform.system() == "Windows":
    # Do Windows setup
    pass
elif platform.system() == "Linux":
    # Do Linux setup
    pass
else:
    # Exit with error
    pass

# Define static variables and parameters dictionaries
conversion_types = [
    {
        "type": "CSV to Excel (xlsx)",
        "sym": "csv-xlsx",
    },
    {
        "type": "CSV to Excel (xls)",
        "sym": "csv-xls",
    },
    {
        "type": "Excel (xlsx) to CSV",
        "sym": "xlsx-csv",
    },
    {
        "type": "Excel (xls) to CSV",
        "sym": "xlsx-csv",
    },
]

separators = [
    {"type": "Comma (,)", "sym": ","},
    {"type": "Semicolon (;)", "sym": ";"},
    {"type": "Pipe (|)", "sym": "|"},
    {"type": "Tab (-->)", "sym": "\t"},
]

decimals = [{"type": "Period (.)", "sym": "."}, {"type": "Comma (,)", "sym": ","}]

quotechars = [
    {"type": "Single Quotes (')", "sym": "'"},
    {"type": 'Double Quotes (")', "sym": '"'},
]

encodings = [
    {"type": "UTF-8", "sym": "utf-8"},
    {"type": "UTF-16", "sym": "utf-16"},
    {"type": "Latin", "sym": "latin_1"},
    {"type": "ASCII", "sym": "ascii"},
]

# Define errors and messages
ERROR_0 = "Error!"
ERROR_1 = "Error code: 001\nPlease select one or more files to be converted."
ERROR_2 = "Error code: 002\nSelected files are not in line with the conversion mode."
ERROR_3 = "Error code: 003\nPlease select a destination folder."
RESULTS_0 = "Conversion executed!"
RESULTS_1 = "All files were successfully converted"
RESULTS_2 = "The following files were not converted due to formatting errors:\n{}\n\nPlease check the conversion parameters."

# Define Menu and versioning
VERSION = "Version 0.1 - 2021-10-11"
CONTACTS = "For support and inquiries, please visit https://gg8.eu/ or send an email to mail@provider.com"
MENU = [["&File", ["&Info", "&Exit"]]]

# Set appearence settings
sg.theme("LightGrey")
sg.theme_input_background_color(color="white")
sg.theme_input_text_color(color="black")
sg.theme_background_color(color="gray80")
sg.theme_text_element_background_color(color="gray80")
font = ("Arial, 14")
sg.set_options(font=font)
# Logo needs to be implemented
# LOGO = get_logo()
# sg.set_global_icon(LOGO)


def set_display_files_list(files_list):
    """
    Function to format the files selected for conversion
    creating a shorter list and total size indication
    when more than 3 files are selected
    """
    if len(files_list) <=3:
        display_files_list = files_list
    else:
        display_files_list = files_list[:3]
        display_files_list.append("...")
        display_files_list.append("and {} more files.".format(len(files_list) - 3))
    return display_files_list


def get_parameters(values, conversion_dict):
    """
    Function to get conversion type and parameters
    in order to perform file conversion.
    Return a dictionary with all parameters
    """
    params = {}
    params["origin_ext"] = conversion_dict.get("sym").split("-")[0]
    params["target_ext"] = conversion_dict.get("sym").split("-")[1]
    separator_dict = next(
        (
            item
            for item in separators
            if item.get("type", "NA") == values.get("separator")[0]
        ),
        None,
    )
    decimal_dict = next(
        (
            item
            for item in decimals
            if item.get("type", "NA") == values.get("decimal")[0]
        ),
        None,
    )
    quotechar_dict = next(
        (
            item
            for item in quotechars
            if item.get("type", "NA") == values.get("quotechar")[0]
        ),
        None,
    )
    encoding_dict = next(
        (
            item
            for item in encodings
            if item.get("type", "NA") == values.get("encoding")[0]
        ),
        None,
    )
    params["separator"] = separator_dict.get("sym")
    params["decimal"] = decimal_dict.get("sym")
    params["quotechar"] = quotechar_dict.get("sym")
    params["encoding"] = encoding_dict.get("sym")
    return params


def convert_file(f, destination_folder, params):
    """
    Utility function to convert csv files to excel
    or excel to csv
    Returns true upon successful conversion
    else return false
    """
    orig_filename = f.rsplit(os.path.sep, 1)[1]
    dest_filename = "{0}.{1}".format(
        orig_filename.rsplit(".", 1)[0], params.get("target_ext")
    )
    dest_path = os.path.join(destination_folder, dest_filename)
    try:
        if params.get("origin_ext") == "csv":
            temp_df = pd.read_csv(
                f,
                sep=params.get("separator"),
                decimal=params.get("decimal"),
                quotechar=params.get("quotechar"),
                encoding=params.get("encoding"),
            )
            temp_df.to_excel(dest_path)
        else:
            temp_df = pd.read_excel(f)
            temp_df.to_csv(
                dest_path,
                sep=params.get("separator"),
                decimal=params.get("decimal"),
                quotechar=params.get("quotechar"),
                encoding=params.get("encoding"),
            )
        return True
    except:
        return False


def main_window():
    """
    Function to create the main program window
    to select files, conversion type and parameters,
    and destination folder
    """
    layout = [
        [sg.Menu(MENU)],
        [sg.T("Select conversion type")],
        [
            sg.Listbox(
                values=[conv.get("type", "NA") for conv in conversion_types],
                size=(30, 4),
                select_mode="LISTBOX_SELECT_MODE_SINGLE",
                enable_events=True,
                auto_size_text=True,
                no_scrollbar=True,
                key="conversion_type",
            )
        ],
        [
            sg.FilesBrowse(
                button_text="Select file(s) to convert",
                disabled=True,
                enable_events=True,
                key="select_files",
                files_delimiter=";",
            )
        ],
        [sg.T("", key="selected_files")],
        [sg.T("Select CSV value separator")],
        [
            sg.Listbox(
                values=[sep.get("type", "NA") for sep in separators],
                default_values=[
                    separators[0].get("type", "NA"),
                ],
                size=(30, 4),
                auto_size_text=True,
                no_scrollbar=True,
                key="separator",
            )
        ],
        [sg.T("Select CSV decimal notation")],
        [
            sg.Listbox(
                values=[dec.get("type", "NA") for dec in decimals],
                default_values=[
                    decimals[0].get("type", "NA"),
                ],
                size=(30, 2),
                select_mode="LISTBOX_SELECT_MODE_SINGLE",
                auto_size_text=True,
                no_scrollbar=True,
                key="decimal",
            )
        ],
        [sg.T("Select CSV field quotation character")],
        [
            sg.Listbox(
                values=[quote.get("type", "NA") for quote in quotechars],
                default_values=[
                    quotechars[0].get("type", "NA"),
                ],
                size=(30, 2),
                select_mode="LISTBOX_SELECT_MODE_SINGLE",
                auto_size_text=True,
                no_scrollbar=True,
                key="quotechar",
            )
        ],
        [sg.T("Select CSV encoding")],
        [
            sg.Listbox(
                values=[enc.get("type", "NA") for enc in encodings],
                default_values=[
                    encodings[0].get("type", "NA"),
                ],
                size=(30, 4),
                select_mode="LISTBOX_SELECT_MODE_SINGLE",
                auto_size_text=True,
                no_scrollbar=True,
                key="encoding",
            )
        ],
        [
            sg.FolderBrowse(
                button_text="Select destination folder",
                disabled=True,
                key="destination_folder",
            )
        ],
        [sg.HorizontalSeparator()],
        [sg.Exit(), sg.T(" "*50), sg.OK()],
    ]
    window = sg.Window("Select Files for conversion", layout, margins=(30,15))
    return window


# Main loop
window = main_window()
while True:
    event, values = window.Read()
    print(event, values)
    if event is None or event == "Exit":
        break
    # Display info
    if event == "Info":
        sg.popup("Info", "\n\n".join([VERSION, CONTACTS]))
    # Enable files and options selections
    # when conversion type has been selected
    if event == "conversion_type":
        conversion_dict = next(
            (
                item
                for item in conversion_types
                if item.get("type", "NA") == values.get("conversion_type")[0]
            ),
            None,
        )
        window.Element("select_files").Update(disabled=False)
        window.Element("destination_folder").Update(disabled=False)
    if event == "select_files":
        files_list = [
            f
            for f in values.get("select_files").split(";")
            if f.endswith(conversion_dict.get("sym").split("-")[0])
        ]
        display_files_list = set_display_files_list(files_list)
        window.Element("selected_files").Update(
            "Selected files:\n{}".format("\n".join(display_files_list))
        )
    if event == "OK":
        # Return an error if no files have been selected
        if values.get("select_files") == "":
            sg.popup(ERROR_0, ERROR_1)
        else:
            if files_list == []:
                sg.popup(ERROR_0, ERROR_2)
            else:
                if values.get("destination_folder") == "":
                    sg.popup(ERROR_0, ERROR_3)
                else:
                    params = get_parameters(values, conversion_dict)
                    results_list = []
                    for i, f in enumerate(files_list):
                        if not sg.one_line_progress_meter(
                            "File conversion progress",
                            i,
                            len(files_list),
                            orientation="h",
                        ):
                            break
                        results_list.append(
                            convert_file(f, values.get("destination_folder"), params)
                        )
                    sg.one_line_progress_meter_cancel()
                    # If all files are successfully converted, return success popup
                    if all(results_list):
                        sg.popup(RESULTS_0, RESULTS_1)
                    # Else flag the failed files
                    else:
                        failed_files = [
                            f for f, k in zip(files_list, results_list) if k == False
                        ]
                        failed_files_string = "\n".join(failed_files)
                        sg.popup(RESULTS_0, RESULTS_2.format(failed_files_string))

# Close window upon exit
window.Close()
