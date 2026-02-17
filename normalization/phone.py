import re


def normalize_phone(phone: str | int) -> str | None:
    if phone is None:
        return None

    phone = str(phone)
    digits = re.sub(r"\D", "", phone)

    if len(digits) >= 10:
        return digits[-10:]

    return None
