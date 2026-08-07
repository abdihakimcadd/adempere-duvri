import os

DATA_DIR = "./data"

def load_md_file(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return f.read()
    except FileNotFoundError:
        return f"[ERROR: {filepath} not found.]"

def get_context_block():
    # Load all 6 DUVRI files
    files_to_load = [
        "dlgs_81_08_art26.md",
        "duvri_mandatory_contents.md",
        "duvri_structure.md",
        "risk_matrices.md",
        "interference_logic.md",
        "user_input_schema.md"
    ]
    
    context = ""
    for filename in files_to_load:
        filepath = os.path.join(DATA_DIR, filename)
        context += f"\n\n--- {filename} ---\n{load_md_file(filepath)}\n"
        
    return context