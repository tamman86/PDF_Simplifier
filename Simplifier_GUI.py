import FreeSimpleGUI as sg
from Simpifier_Processing import extract_text_from_pdf, clean_text, save_text_to_word

sg.theme("SystemDefault")

layout = [
    [sg.Text("Select PDF(s) to convert:")],
    [sg.Input(key="-FILES-", readonly=True, enable_events=True),
     sg.FilesBrowse("Browse Files", file_types=(("PDF Files", "*.pdf"),)),
     sg.Button(button_text="Clear", key="-CLEAR-")],
    [sg.Checkbox("Remove Newline", key="-NEWLINE-"),sg.Checkbox("Remove Hyphenation", key="-HYPHENATION-")],
    [sg.Checkbox("Remove Double Spacing", key="-DOUBLESPACE-"),
     sg.Checkbox("Restore Paragraph", key="-PARAGRAPH-")],
    [sg.Button("Convert PDF File", key="-CONVERT-"), sg.Button("Exit")]
]

window = sg.Window("PDF to Word Converter", layout,
                   finalize=True)

while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED or event == "Exit":
        break

    elif event == "-CLEAR-":
        window["-FILES-"].update("")

    elif event == "-CONVERT-":

        newline = values["-NEWLINE-"]
        hyphen = values["-HYPHENATION-"]
        doublespace = values["-DOUBLESPACE-"]
        paragraph = values["-PARAGRAPH-"]

        pdf_paths_str = values["-FILES-"]
        pdf_files = pdf_paths_str.split(';')

        for pdf_path in pdf_files:
            raw_text = extract_text_from_pdf(pdf_path)
            text, suffixes = clean_text(raw_text, newline, hyphen, doublespace, paragraph)

            suffixes += ".docx"
            output_file = pdf_path.replace(".pdf", suffixes)
            save_text_to_word(text, output_file)