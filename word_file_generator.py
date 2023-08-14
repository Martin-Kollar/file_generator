import win32com.client as win32
import os
import sys
import yaml


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


def save_and_close_documents(*docs):
    for doc in docs:
        doc.Save()
        doc.Close()


def main(template_id, macro_id):
    cfg = None
    parent_dir = os.path.dirname(os.path.abspath(__file__))
    config_file = os.path.join(parent_dir, "config.yml")
    template_dir = os.path.join(parent_dir, "templates")
    output_dir = os.path.join(parent_dir, "output")

    with open(config_file, "r") as stream:
        try:
            cfg = yaml.safe_load(stream)
        except yaml.YAMLError as exc:
            print(
                f"Error reading configuration from '{config_file}': {exc}")
            sys.exit(1)

    # Get template and macro paths from config using template_id and macro_id
    template_name = cfg['templates']['word'][template_id]
    macro_name = cfg['macros'][macro_id]

    template_path = os.path.join(template_dir, "files", template_name)
    macro_path = os.path.join(template_dir, "macros", macro_name)

    # Create an instance of Word
    word_app = win32.gencache.EnsureDispatch('Word.Application')
    word_app.Visible = False
    doc = word_app.Documents.Open(template_path)

    WordDoc = win32.gencache.EnsureDispatch(word_app.Documents(1))

    # Insert macro from text file
    insert_macro_from_file(macro_path, word_app)

    output_path = os.path.join(output_dir, os.path.splitext(
        WordDoc.FullName)[0] + "_output.docm")
    # 16 corresponds to .docm format
    WordDoc.SaveAs(output_path, FileFormat=16)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python script.py <template_id> <macro_id>")
        sys.exit(1)

    template_id = int(sys.argv[1])
    macro_id = int(sys.argv[2])

    main(template_id, macro_id)
