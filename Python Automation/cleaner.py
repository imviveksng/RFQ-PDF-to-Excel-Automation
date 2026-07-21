from datetime import datetime


def clean_text(value):

    if not value:
        return ""

    value = " ".join(value.split())

    return value.upper()


def clean_date(value):

    if not value:
        return ""

    try:

        date = datetime.strptime(value, "%d-%b-%Y")

        return date.strftime("%d-%m-%Y")

    except ValueError:

        return value


def get_month(date_string):

    if not date_string:

        return ""

    try:

        date = datetime.strptime(date_string, "%d-%m-%Y")

        return date.strftime("%b-%y").upper()

    except:

        return ""