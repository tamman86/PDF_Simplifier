import FreeSimpleGUI as sg
from Simpifier_Processing import extract_text_from_pdf, clean_text, save_text_to_word

sg.theme("SystemDefault")

layout = [
    [sg.Text("Select PDF(s) to convert:")],
    [sg.Input(key="-FILES-", readonly=True, enable_events=True),
     sg.FilesBrowse("Browse Files", file_types=(("PDF Files", "*.pdf"),)),
     sg.Button(button_text="Clear", key="-CLEAR-")],
    [sg.Text("Extraction Mode (Select One): ")],
    [
        sg.Checkbox("Generic", key='-GENERIC-', enable_events=True, default=True),
        sg.Checkbox("PDF Plumber", key='-PDFPLUMBER-', enable_events=True, default=False),
        sg.Checkbox("OCR", key='-OCR-', enable_events=True, default=False)
    ],
    [sg.Text("Cleaning Techniques : ")],
    [sg.Checkbox("Remove Newline", key="-NEWLINE-"),sg.Checkbox("Remove Hyphenation", key="-HYPHENATION-")],
    [sg.Checkbox("Remove Double Spacing", key="-DOUBLESPACE-"),
     sg.Checkbox("Restore Paragraph", key="-PARAGRAPH-")],
    [sg.Button("Convert PDF File", key="-CONVERT-"), sg.Button("Exit")]
]

window = sg.Window("PDF to Word Converter", layout,
                   finalize=True)

# Store the keys of the checkboxes for easy iteration
checkbox_keys = ['-GENERIC-', '-PDFPLUMBER-', '-OCR-']
currently_selected_key = None

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
        use_generic = values["-GENERIC-"]
        use_pdfplumber = values["-PDFPLUMBER-"]
        use_ocr = values["-OCR-"]

        pdf_paths_str = values["-FILES-"]
        pdf_files = pdf_paths_str.split(';')

        for pdf_path in pdf_files:
            raw_text, suffixes = extract_text_from_pdf(pdf_path, use_generic, use_pdfplumber, use_ocr)
            text, suffixes = clean_text(raw_text, newline, hyphen, doublespace, paragraph, suffixes)

            suffixes += ".docx"
            output_file = pdf_path.replace(".pdf", suffixes)
            save_text_to_word(text, output_file)

    elif event in checkbox_keys:
        # The checkbox that was just clicked
        clicked_key = event

        # If the clicked checkbox is now True (checked)
        if values[clicked_key]:
            # Uncheck all other checkboxes
            for key in checkbox_keys:
                if key != clicked_key:
                    window[key].update(value=False)
            currently_selected_key = clicked_key
        else:
            # If the user unchecks the currently selected checkbox,
            # we don't want to automatically re-check it.
            # So, if the clicked_key was the currently_selected_key and it's now False,
            # update currently_selected_key to None.
            if clicked_key == currently_selected_key:
                currently_selected_key = None