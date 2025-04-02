import os


def valid_xlsx(path: str, exists=True):
    allowed_exts = (".xlsx", ".xls")

    if exists and not os.path.exists(path):
        raise FileExistsError(f"File {path} doesn't exists!")
    
    elif not any(path.endswith(ext) for ext in allowed_exts):
        raise ValueError(f"Allowed extensions: {allowed_exts}")
