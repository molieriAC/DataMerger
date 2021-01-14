import fnmatch
import os
import pandas as pd
import pyreadstat as prd
import PySimpleGUI as sg


def FileFinder(d, k):
    """

    Args:
        d: string -- indicates the folder that files should be in
        k: string -- part of the file name. Indicates which files should be added the the return list.

    Returns: list of files

    """
    # File Finder: Open and find all excel files. Will iterate over all files in values["-DIR-"].
    savs = []
    for root, dirs, files in os.walk(d):
        for f in files:
            if fnmatch.fnmatch(f, "*" + k + "*"):
                savs.append(os.path.join(root, f))
    return savs


def MergeRunner(inputs, files):
    """Merges the files that are passed in based on informaion from the GUI

    Args:
        inputs: determines which type of merge should be done depending on which file type was selected by the GUI,
        and directs the merge operations

        files: A list of files to be merged

    Returns: outputs the file location to console

    """
    # File Merger for spss files
    if inputs["-FileType-"] == "spss":
        filesList = []
        valLabels = {}
        # missingLabels = {}
        for i in files:
            newF, spssMeta = prd.read_sav(i, user_missing=True, disable_datetime_conversion=True)
            filesList.append(newF)
            # Loop through the value labels dictionary to update the sub-dicts
            if not bool(valLabels):
                valLabels.update(spssMeta.variable_value_labels)
            else:
                for j in valLabels:
                    valLabels[j].update(spssMeta.variable_value_labels[j])
            # missingLabels.update(spssMeta.missing_ranges)
        allFiles = pd.concat(filesList, sort=False)
        prd.write_sav(allFiles, inputs["-DIR-"] + "/" + inputs["-OutFile-"] + ".sav",
                      variable_value_labels=valLabels)
        return "merged spss file saved at: " + inputs["-DIR-"] + "/" + inputs["-OutFile-"] + ".sav"

    # File Merger for excel files
    if inputs["-FileType-"] == "excel":
        filesList = []
        for i in files:
            filesList.append(pd.read_excel(i, sheet_name=inputs["-ExcelSheet-"]))
        allFiles = pd.concat(filesList, sort=False)
        allFiles.to_excel(inputs["-DIR-"] + "/" + inputs["-OutFile-"] + ".xlsx", index=False)
        return "merged excel file saved at: " + inputs["-DIR-"] + "/" + inputs["-OutFile-"] + ".xlsx"


def main():
    """Runs everything"""
    # Set up the GUI
    sg.theme('DarkGrey4')
    # All the stuff inside your window.
    layout = [[sg.Text('Select the root directory'), sg.FolderBrowse(key="-DIR-"),
               sg.Text('Select the file type'), sg.Combo(["excel", "spss"], key="-FileType-")],
              [sg.Text("Text used for matching files (case sensitive)"), sg.InputText(key="-MatchKey-")],
              [sg.Text('If excel, what is the sheet name? (case sensitive)'), sg.InputText(key="-ExcelSheet-")],
              [sg.Text('Name the output file (no extension needed)'), sg.InputText(key="-OutFile-")],
              [sg.Button('Ok'), sg.Button('Cancel')]]

    # Create the Window
    window = sg.Window('Window Title', layout)
    # Event Loop to process "events" and get the "values" of the inputs
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancel':  # if user closes window or clicks cancel
        exit()
    window.close()

    # get the list of files to be merged
    saveList = FileFinder(values["-DIR-"], values["-MatchKey-"])
    # run the merge and output file location to the console
    merger = MergeRunner(values, saveList)
    print(merger)


if __name__ == '__main__':
    main()
