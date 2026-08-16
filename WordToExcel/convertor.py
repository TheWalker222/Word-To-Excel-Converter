from datetime import datetime

def convertValue(value):
    value = value.strip()
    if "$" in str(value):
        method = "money$"
        value = value.replace("$", "")
        value = value.replace(",", ".")
        try:
            return float(value), method
        except ValueError:
            pass
    if "€" in str(value):
            method = "money€"
            value = value.replace("€", "")
            value = value.replace(",", ".")
            try:
                return float(value), method
            except ValueError:
                pass
    try:
        method = "dateDDMMYYYY"
        return datetime.strptime(value, "%d.%m.%Y"), method
    except ValueError:
        pass
    try:
        return int(value), None
    except ValueError:
        pass
    try:
        return float(value), None
    except ValueError:
        pass
    return value, None