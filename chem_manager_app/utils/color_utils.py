def hex_to_bgr(hex_str):
    """
    Converts a hex color string like '#RRGGBB' to Excel's BGR integer format.
    """
    if not hex_str or not hex_str.startswith("#") or len(hex_str) != 7:
        return 16777215 # White default
    
    r = int(hex_str[1:3], 16)
    g = int(hex_str[3:5], 16)
    b = int(hex_str[5:7], 16)
    
    return (b << 16) + (g << 8) + r

def bgr_to_hex(bgr_int):
    """
    Converts an Excel BGR integer to a hex color string '#RRGGBB'.
    """
    b = (bgr_int >> 16) & 0xFF
    g = (bgr_int >> 8) & 0xFF
    r = bgr_int & 0xFF
    
    return f"#{r:02x}{g:02x}{b:02x}"
