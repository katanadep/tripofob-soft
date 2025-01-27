import os
import re
import zipfile
import xml.etree.ElementTree as ET
from rich.console import Console
from rich.panel import Panel
from rich.text import Text

console = Console()

def search_in_file(file_path, search_pattern):
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
            lines = file.readlines()
            found_lines = []
            for line in lines:
                if re.search(search_pattern, line, re.IGNORECASE):
                    found_lines.append(line.strip())
            if found_lines:
                console.print(Panel(Text("\n".join(found_lines), style="bold green"), title=f"Found in {file_path}", border_style="green"))

    except Exception as e:
        console.print(f"Error reading {file_path}: {e}", style="red")

def search_in_xlsx_file(file_path, search_pattern):
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            for name in z.namelist():
                if name.startswith('xl/worksheets/sheet'):
                    with z.open(name) as sheet:
                        tree = ET.parse(sheet)
                        root = tree.getroot()
                        found_values = []
                        for row in root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
                            for cell in row.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v'):
                                value = cell.text
                                if value and re.search(search_pattern, value, re.IGNORECASE):
                                    found_values.append(value)
                        if found_values:
                            console.print(Panel(Text("\n".join(found_values), style="bold white on black"), title=f"Found in {file_path}", border_style="green"))

    except Exception as e:
        console.print(f"Error reading {file_path}: {e}", style="red")

def search_in_directory(directory, search_pattern):
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(('.xls', '.db', '.sql', '.txt', '.csv', '.json', '.sqlite', '.xml', '.html', '.xlsx')):
                file_path = os.path.join(root, file)
                if file.endswith('.xlsx'):
                    search_in_xlsx_file(file_path, search_pattern)
                else:
                    search_in_file(file_path, search_pattern)

if __name__ == "__main__":
    banner = """
    ████████╗██████╗░██╗██████╗░░█████╗░███████╗░█████╗░██████╗░
    ╚══██╔══╝██╔══██╗██║██╔══██╗██╔══██╗██╔════╝██╔══██╗██╔══██╗
    ░░░██║░░░██████╔╝██║██████╔╝██║░░██║█████╗░░██║░░██║██████╦╝
    ░░░██║░░░██╔══██╗██║██╔═══╝░██║░░██║██╔══╝░░██║░░██║██╔══██╗
    ░░░██║░░░██║░░██║██║██║░░░░░╚█████╔╝██║░░░░░╚█████╔╝██████╦╝
    ░░░╚═╝░░░╚═╝░░╚═╝╚═╝╚═╝░░░░░░╚════╝░╚═╝░░░░░░╚════╝░╚═════╝░
    """
    console.print(Panel(banner, title="TRIPOFOB", style="bold white on black", expand=False))
    print("")

    directory_to_search = os.path.dirname(os.path.abspath(__file__))
    search_pattern = console.input("Введите данные для поиска: ")
    print("")
    print("")

    search_in_directory(directory_to_search, search_pattern)
    console.input("Нажмите Enter для выхода...")