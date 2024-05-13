from sys import argv, platform
from os import urandom, mkdir, walk, path, getlogin
from zipfile import ZipFile, BadZipFile
from shutil import rmtree

# Checking command syntax
doc_path = argv[-1]
if len(argv) != 2 or not doc_path.endswith(".docx"):
    print("Syntax: python3 analyze.py /path/to/doc/name.docx")
    exit(1)

# Setting the temp directory based on the OS
# If your OS is unusual, edit this accordingly!
if platform == "win32":
    tmp_path = f"C:/Users/{getlogin()}/AppData/Local/Temp"
else:
    tmp_path = "/tmp"

# Creating the temporary directory of the form 'temp_<random_hex>'
zip_path = path.join(tmp_path, f"temp_{urandom(16).hex()}")
mkdir(zip_path)

# Checking that the file entered exists and is ZIP compressed
# Then unzipping the file to the temporary directory
try:
    with ZipFile(doc_path, "r") as zf:
        zf.extractall(zip_path)
except (BadZipFile, FileNotFoundError):
    print(f"Error: '{doc_path}' is not a Word document or does not exist.")
    rmtree(zip_path)
    exit(1)

print(f"*** Analyzing rels from '{doc_path.split('/')[-1]}' ***")

# Walking through the files of the decompressed Word document
# And getting paths for files with a '.rels' extension
rels_paths = []
for root, _, files in walk(zip_path):
    for file in files:
        if file.endswith(".rels"):
            rels_paths.append(path.join(root, file))

# Parsing through '.rels' files to find all targets used
for rels_path in rels_paths:
    with open(rels_path, "r") as f:
        print(f"Targets from {rels_path.replace(zip_path + '/', '')}:")
        for target in f.read().split("Target=\"")[1:]:
            print("\t", target.split("\"/>")[0])

    print()

rmtree(zip_path)
