import json
import sys
from pathlib import Path
from tkinter import filedialog, messagebox
from tkinter.filedialog import askopenfilename
from typing import Dict, List, Any
import pandas as pd


def main():
    try:
        current_dir_path = get_base_dir()
        
        file_path = get_file_path()
        if not file_path:
            return
            
        data_dict = get_raw_dict(file_path=file_path)
        if data_dict is None:
            return
            
        json_path = get_json_path(start_dir=current_dir_path)
        if not json_path:
            return
        
        success = save_json_data(data_dict=data_dict, path=json_path)
        if not success:
            return
        messagebox.showinfo(
            title="Success",
            message=f"Data successfully saved to:\n{json_path}"
        )
    except Exception as main_e:
        messagebox.showerror(
            title="Application Error",
            message=f"An unexpected critical error occurred:\n{type(main_e).__name__}: {main_e}"
        )

def get_base_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    else:
        return Path(__file__).resolve().parent

def get_file_path() -> Path | None:
    raw_path = askopenfilename(
        title="Choose Excel Files", filetypes=[("Excel Documents", "*.xlsx")]
    )
    return Path(raw_path) if raw_path else None


def get_raw_dict(*, file_path: Path) -> Dict[str, List[Any]] | None:
    try:
        df = pd.read_excel(file_path)

        if df.empty:
            messagebox.showerror(title="File is empty", message="Excel file contains no data rows.")
            return None
            
        raw_dict = df.to_dict(orient="list")
        return dict_cleanup(raw_dict=raw_dict)
        
    except FileNotFoundError:
        messagebox.showerror(title="File not found", message="Excel file not found. Please check if the file exists.")
        return None
    except ValueError as e: 
        if "No sheets" in str(e) or "Worksheet" in str(e):
            messagebox.showerror(title="File is empty", message="Excel file is completely empty or corrupted.")
        else:
            messagebox.showerror(title="Unexpected error", message=f"Data error: {e}")
        return None
    except Exception as e:
        messagebox.showerror(title="Unexpected error", message=f"An unexpected error occurred while processing:\n{e}")
        return None


def dict_cleanup(raw_dict: Dict[str, List[Any]]) -> Dict[str, List[Any]]:
    cleaned_dict = {}
    if not raw_dict:
        return cleaned_dict
        
    for key, values in raw_dict.items():
        key_str = str(key).strip()
        if not key_str or key_str.startswith("Unnamed:"):
            continue
            
        cleaned_values = []
        for val in values:
            if pd.isna(val):
                continue
            if isinstance(val, str):
                val = val.strip()
                if not val:
                    continue
            cleaned_values.append(val)
            
        if cleaned_values:
            cleaned_dict[key_str] = cleaned_values
            
    return cleaned_dict


def get_json_path(*, start_dir: Path) -> Path | None:
    ext: str = ".json"
    default_name: str = "data"
    
    raw_path = filedialog.asksaveasfilename(
        initialdir=start_dir,
        initialfile=f"{default_name}{ext}",
        title="Choose folder and select JSON-file name",
        defaultextension=ext,
        filetypes=[("Files", f"*{ext}"), ("All files", "*.*")]
    )
    
    if not raw_path:
        return None
        
    json_path = Path(raw_path)
    if json_path.suffix.lower() != ext:
        json_path = json_path.with_suffix(ext)
    return json_path


def save_json_data(*, data_dict: Dict[str, List[Any]], path: Path)->bool:
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data_dict, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        error_title = "Save Error"
        error_message = (
            f"Failed to save data to the JSON file.\n\n"
            f"Path: {path.name}\n"
            f"Please check if the file is open in another program or if you have write permissions.\n\n"
            f"Details: {type(e).__name__}: {e}"
        )
        messagebox.showerror(title=error_title, message=error_message)
        return False

if __name__ == "__main__":
    main()
