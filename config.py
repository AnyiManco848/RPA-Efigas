import os
from dotenv import load_dotenv

load_dotenv()

EXCEL_PATH = os.getenv("EXCEL_PATH", "Contratos Efigas.xlsx")
PDFS_FOLDER = os.getenv("PDFS_FOLDER", "Pdfs")


def get_credentials(email: str) -> tuple[str, str]:
    """Devuelve (email, password) buscando entre USER1_, USER2_, ... en .env.
    Lanza ValueError si el email no tiene credenciales configuradas."""
    i = 1
    while True:
        env_email = os.getenv(f"USER{i}_EMAIL")
        if env_email is None:
            break
        if env_email.strip().lower() == email.strip().lower():
            password = os.getenv(f"USER{i}_PASSWORD", "")
            return env_email.strip(), password
        i += 1
    raise ValueError(f"Sin credenciales en .env para: {email}")
