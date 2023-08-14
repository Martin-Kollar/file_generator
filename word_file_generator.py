import win32com.client as win32
import os
import sys
import yaml

cfg = None
parent_dir = os.path.dirname(os.path.abspath(__file__))
config_file = os.path.join(parent_dir, "config.yml")
template_dir = os.path.join(parent_dir, "templates")
output_dir = os.path.join(parent_dir, "output")

with open(config_file, "r") as stream:
    try:
        cfg = yaml.safe_load(stream)
    except yaml.YAMLError as exc:
        print(f"Error reading configuration from '{config_file}': {exc}")
        sys.exit(1)

def load_word_template(template_path):
    app = win32.gencache.EnsureDispatch('Word.Application')
    app.Visible = False
    doc = app.Documents.Open(template_path)
    return app, doc

def insert_macro_from_file(macro_path, target_app):
    with open(macro_path, 'r') as macro_file:
        macro_content = macro_file.read()

    target_macros = target_app.VBE.VBProjects(1).VBComponents
    new_macro = target_macros.Add(1)
    new_macro.CodeModule.AddFromString(macro_content)

def main(template_id, macro_id):
    # Get template path from config using template_id
    template_name = cfg['templates']['word'][template_id]
    template_path = os.path.join(template_dir, "files", template_name)

    # Load the Word template (template with the associated macro)
    word_app, template_doc = load_word_template(template_path)

    # Create a new document based on the template
    new_doc = template_doc.NewDocument()
    template_doc.Close()

    # Save the new document in the output folder
    new_doc_path = os.path.join(output_dir, "new_document.docx")
    new_doc.SaveAs(new_doc_path)

    # Close the new document
    new_doc.Close()

    # Re-open the newly created document
    new_doc = word_app.Documents.Open(new_doc_path)

    # Insert macro from text file into the new document
    macro_name = cfg['macros'][macro_id]
    macro_path = os.path.join(template_dir, "macros", macro_name)
    insert_macro_from_file(macro_path, new_doc.Application)

    # Save the output document with the macro inserted
    output_filename = os.path.splitext(os.path.basename(template_name))[0]
    output_path = os.path.join(output_dir, output_filename)
    new_doc.SaveAs(output_path + "_output.docm", FileFormat=16)

    # Close the documents
    new_doc.Close()
    word_app.Quit()

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python script.py <template_id> <macro_id>")
        sys.exit(1)

    template_id = int(sys.argv[1])
    macro_id = int(sys.argv[2])

    main(template_id, macro_id)
