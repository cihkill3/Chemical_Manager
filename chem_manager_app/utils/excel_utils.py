def get_col_letter(col_idx):
    """Convert 1-based column index to Excel column letter (e.g. 1 -> A, 28 -> AB)"""
    letter = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter
