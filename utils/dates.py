# utils/dates.py

from datetime import datetime

def format_month_label(date_obj):
    """
    Needs datetime-object to create name 'Jan26'.
    """
    if not isinstance(date_obj, datetime):
        raise TypeError("format_month_label needs date format")

    return date_obj.strftime("%b%y")


def format_month_label_german(date_obj):
    """
    Creates German month abbreviation from datetime object.
    Returns format like 'Jan', 'Feb', 'Mrz', etc.
    
    Args:
        date_obj: datetime object
        
    Returns:
        str: German month abbreviation (e.g., 'Jan', 'Feb', 'Mrz')
    """
    if not isinstance(date_obj, datetime):
        raise TypeError("format_month_label_german needs datetime format")
    
    # Mapping of month numbers to German abbreviations
    german_months = {
        1: "Jan",
        2: "Feb",
        3: "Mrz",
        4: "Apr",
        5: "Mai",
        6: "Jun",
        7: "Jul",
        8: "Aug",
        9: "Sep",
        10: "Okt",
        11: "Nov",
        12: "Dez"
    }
    
    return german_months[date_obj.month]
