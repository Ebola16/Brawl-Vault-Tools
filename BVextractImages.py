import os
import re
import requests
import pandas as pd
from urllib.parse import urlparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock

# ------------------------------
# Config
# ------------------------------

MAX_WORKERS = 10
MIN_VALID_IMAGE_SIZE = 25000  # filters invalid Imgur placeholders

# ------------------------------
# Helpers
# ------------------------------

def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "", name).strip()


def get_extension_from_url(url: str) -> str:
    parsed = urlparse(url)
    _, ext = os.path.splitext(parsed.path)
    return ext if ext else ".jpg"


def create_session():
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Referer": "https://imgur.com/"
    })
    return session


def download_image(session, link, filepath, title, failed_titles, lock):
    try:
        response = session.get(link, timeout=20)
        response.raise_for_status()

        content_type = response.headers.get("Content-Type", "")
        if "image" not in content_type:
            raise Exception("Not an image")

        content = response.content

        # Detect invalid Imgur placeholder image
        if "imgur.com" in link.lower() and len(content) < MIN_VALID_IMAGE_SIZE:
            raise Exception("Likely invalid Imgur image")

        with open(filepath, "wb") as f:
            f.write(content)

        print(f"Downloaded: {os.path.basename(filepath)}")

    except Exception:
        with lock:
            failed_titles.add(title)
        print(f"Failed (Title flagged): {title}")


# ------------------------------
# Main
# ------------------------------

def main():
    script_folder = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(script_folder, "brawl_vault_full_export.xlsx")
    images_folder = os.path.join(script_folder, "images")
    missing_log_path = os.path.join(script_folder, "Missing Images.txt")

    if not os.path.exists(excel_path):
        print("brawl_vault_full_export.xlsx not found.")
        return

    os.makedirs(images_folder, exist_ok=True)

    df = pd.read_excel(excel_path, engine="openpyxl")

    if "Image URLs" not in df.columns:
        print("Column 'Image URLs' not found.")
        return

    failed_titles = set()
    lock = Lock()
    tasks = []

    session = create_session()

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:

        for _, row in df.iterrows():
            title_raw = str(row.get("Title", "Untitled")).strip()
            title = sanitize_filename(title_raw)

            links_cell = row.get("Image URLs")

            if pd.isna(links_cell):
                continue

            links = [l.strip() for l in str(links_cell).split(",") if l.strip()]

            for index, link in enumerate(links, start=1):
                extension = get_extension_from_url(link)
                filename = f"{title}_{index}{extension}"
                filepath = os.path.join(images_folder, filename)

                if os.path.exists(filepath):
                    continue

                tasks.append(
                    executor.submit(
                        download_image,
                        session,
                        link,
                        filepath,
                        title_raw,  # store original title in log
                        failed_titles,
                        lock
                    )
                )

        for future in as_completed(tasks):
            pass

    # Write failed titles
    if failed_titles:
        with open(missing_log_path, "w", encoding="utf-8") as f:
            for title in sorted(failed_titles):
                f.write(title + "\n")

        print("\nMissing Images.txt created (titles with failed images).")
    else:
        print("\nAll images downloaded successfully.")


if __name__ == "__main__":
    main()
