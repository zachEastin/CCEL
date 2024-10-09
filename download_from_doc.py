from bs4 import BeautifulSoup, Tag
from pathlib import Path
import requests
import re

import subprocess


def sanitize_filename(filename: str, replacement: str = "_") -> str:
    """
    Sanitize a string to be safe for use as a filename in a file system path.

    Parameters:
    - filename: The original string to sanitize.
    - replacement: The string used to replace invalid characters (default is '_').

    Returns:
    - A sanitized string safe to use as a filename.
    """
    # Define invalid characters based on cross-platform restrictions
    # This includes reserved characters like / \ ? % * : | " < > — and non-printable control characters
    invalid_characters = (
        r'[<>:"/\\|—?*\x00-\x1F]'  # Windows-invalid characters and control characters
    )

    # Substitute invalid characters with the replacement string
    sanitized = re.sub(invalid_characters, replacement, filename)

    # Optionally strip leading/trailing spaces, if necessary
    sanitized = sanitized.strip()

    # Avoid reserved file names (Windows-specific)
    reserved_names = {
        "CON",
        "PRN",
        "AUX",
        "NUL",
        "COM1",
        "COM2",
        "COM3",
        "COM4",
        "COM5",
        "COM6",
        "COM7",
        "COM8",
        "COM9",
        "LPT1",
        "LPT2",
        "LPT3",
        "LPT4",
        "LPT5",
        "LPT6",
        "LPT7",
        "LPT8",
        "LPT9",
    }
    if sanitized.upper() in reserved_names:
        sanitized = f"{sanitized}_{replacement}"

    # Optionally enforce a max length (255 is typical)
    return sanitized[:255]  # Adjust length limit if needed


base_url = "https://ccel.org/"


def parse_html(file_path, force_redownload=False):
    print()
    print()

    with open(file_path, "r", encoding="utf-8") as file:
        soup = BeautifulSoup(file, "xml")

    # Get the tag named "DC.Title" with a "sub" attribute set to "Main"
    title_tag = soup.find("title")
    title = title_tag.text if title_tag else None
    print(f"{title=}")

    result: dict[str, str | list[str, str]] = {}

    p: Tag
    for p in soup.find_all("p", class_="i1"):
        if p.get("id", "").startswith("i-p"):
            hrefs = [
                (a.text, a.get("href"))
                for a in p.find_all("a", recursive=True)
                if a.get("href")
            ]
            if len(hrefs) == 1:
                key = p.find("a").text
                result[key] = hrefs[0][1]
            else:
                key = p.contents[0].strip()
                result[key] = hrefs
        break

    dir_path = Path(__file__).parent / sanitize_filename(title)
    dir_path.mkdir(parents=True, exist_ok=True)
    print(f"{dir_path=}")

    def download_images_in_html_file(url: str, html_file_path: Path, tabs: int):
        with open(html_file_path, "r", encoding="utf-8") as file:
            # soup = BeautifulSoup(file, "html.parser")
            soup = BeautifulSoup(file, "xml")

        img: Tag
        for img in soup.find_all("img"):
            src = img.get("src")
            if not src:
                continue
            tbs = "\t" * tabs
            print(f"{tbs}- {src=}")
            src = src.replace("../", "")

            if not src.startswith("http"):
                src = url.replace(".xml", "/") + src

            img_fp = html_file_path.parent / "files" / Path(src).name
            exists = download_file(src, img_fp, tabs + 1)

            # Set the new image path relative to the HTML file
            if exists:
                img["src"] = img_fp

        with open(html_file_path, "w", encoding="utf-8") as file:
            file.write(soup.decode())

    def download_file(url: str, fp: Path, tabs: int):
        if fp.exists() and not force_redownload:
            tbs = "\t" * tabs
            print(f"{tbs}- Skipped")
            return True

        fp.parent.mkdir(parents=True, exist_ok=True)

        print(f"\t- Downloading: {url}\n\t  -> {fp}")
        with requests.get(url) as r:
            if r.status_code == 200:
                with open(fp, "wb") as f:
                    f.write(r.content)
                    tbs = "\t" * (tabs + 1)
                    print(f"{tbs}- Success")
                return True
            else:
                tbs = "\t" * (tabs + 1)
                print(f"{tbs}- Failed: {r.status_code} | {r.reason}")
                return False

    print()
    for key, value in result.items():
        print(f"Section: {key}")
        if isinstance(value, str):
            if value.endswith(".html"):
                value = value.replace(".html", ".xml")
            fp = (
                dir_path
                / f"{sanitize_filename(key)}___{sanitize_filename(Path(value).name)}"
            )

            url = base_url + value
            exists = download_file(base_url + value, fp, 1)
            if exists and fp.exists() and fp.suffix == ".xml":
                download_images_in_html_file(url, fp, 1)
                html_file = xsl_convert_to_html(fp, url, 1)
                if html_file:
                    html_to_docx(html_file, 1)
        else:
            for name, href in value:
                print(f"\t- Subsection: {name}")
                # Sanitize key for filesystem
                section_path = dir_path / sanitize_filename(key)
                section_path.mkdir(parents=True, exist_ok=True)

                if href.endswith(".html"):
                    href = href.replace(".html", ".xml")

                fp = (
                    section_path
                    / sanitize_filename(name)
                    / sanitize_filename(Path(href).name)
                )
                url = base_url + href
                exists = download_file(url, fp, 2)
                if exists and fp.exists() and fp.suffix == ".xml":
                    download_images_in_html_file(url, fp, 2)
                    html_file = xsl_convert_to_html(fp, url, 2)
                    if html_file:
                        html_to_docx(html_file, 2)

    return result


def xsl_convert_to_html(thml_file: Path, url: str, tabs: int):
    tbs = "\t" * tabs
    print(f"{tbs}Converting ThML to HTML...")

    # Ensure bookInfo.xml exists
    if not (thml_file.parent / "bookInfo.xml").exists():
        url.replace(".xml", "/bookInfo.xml")
        response = requests.get(url)

        if response.status_code == 200:
            with open(thml_file.parent / "bookInfo.xml", "wb") as f:
                f.write(response.content)

    # Load XSLT stylesheet
    xslt_file_src = Path(__file__).parent / "thml.html.xsl"
    xslt_file = thml_file.parent / xslt_file_src.name
    xslt_file.write_bytes(xslt_file_src.read_bytes())
    if not thml_file.exists():
        print(f"{tbs}\t- Failed: ThML file not found: {thml_file}")
        return
    html_file = thml_file.with_suffix(".html")
    args = (
        "java",
        "-jar",
        "C:/SaxonHE/SaxonHE12-5J/saxon-he-12.5.jar",
        f"-s:{str(thml_file)}",
        f"-xsl:{str(xslt_file)}",
        f"-o:{str(html_file)}",
    )

    print(f"{tbs}\t- Running command: {' '.join(args)}")
    proc = subprocess.run(args)

    if proc.returncode != 0:
        print(f"{tbs}\t- Failed: {proc.returncode}")
        return None
    else:
        print(f"{tbs}\t- Success")
        return html_file


def html_to_docx(html_file: Path, tabs: int):
    tbs = "\t" * tabs
    print(f"{tbs}Converting HTML to DOCX...")

    docx_file = html_file.with_suffix(".docx")
    args = (
        "pandoc",
        f"{str(html_file)}",
        "-o",
        f"{str(docx_file)}",
    )

    print(f"{tbs}\t- Running command: {' '.join(args)}")
    proc = subprocess.run(args)

    if proc.returncode != 0:
        print(f"{tbs}\t- Failed: {proc.returncode}")
    else:
        print(f"{tbs}\t- Success")
        return docx_file


# Example usage
file_path = Path(__file__) / "calvins_commentaries.doc"
data = parse_html(file_path)
# print(data)

# c:\Users\jc_4_\Documents\CCEL\Calvin's Commentaries_Complete\Genesis\1-23\calcom01.xml
# C:\Users\jc_4_\Documents\CCEL\Calvin's Commentaries_Complete\Genesis\1-23\calcom01.xml
