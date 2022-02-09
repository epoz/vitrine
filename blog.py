#!/usr/bin/env python
import os, sys, shutil, zlib
import jinja2
from jinja2 import Environment, FileSystemLoader, select_autoescape
from markupsafe import Markup
import pypandoc
import panflute
import io
import requests
import zipfile
from rich import print
from rich.progress import track


DROPBOX_URL = os.getenv("DROPBOX_URL")
TEMPLATE_PATH = os.getenv("TEMPLATE_PATH")
OUT_PATH = os.getenv("OUT_PATH")
EXTRACT_PATH = os.getenv("EXTRACT_PATH")

print(f"Init template environment [green]{TEMPLATE_PATH}[/green]")
env = Environment(
    loader=FileSystemLoader(TEMPLATE_PATH),
    autoescape=select_autoescape(["html", "xml"]),
)


def go(input_path):
    if input_path.endswith(".jpg"):
        print(f"Copying {input_path} to {OUT_PATH}")
        shutil.copy(input_path, OUT_PATH)
    if not input_path.endswith(".docx"):
        return
    outfile_path = input_path.replace(EXTRACT_PATH, OUT_PATH).replace(".docx", ".html")
    outfile_media = input_path.replace(EXTRACT_PATH, OUT_PATH).replace(".docx", "")
    print(input_path, outfile_path, outfile_media)

    html_output = pypandoc.convert_file(
        input_path,
        to="html5",
        extra_args=[
            f"--extract-media={outfile_media}",
            "--data-dir=.",
            "--template=html5template",
        ],
    )
    try:
        template = env.get_template(f"post.html")
    except jinja2.exceptions.TemplateNotFound:
        print("Oops, where is your template?")
        sys.exit(1)

    out = template.render({"content": Markup(html_output)})
    out2 = out.replace(
        'href="!', 'class="btn" style="background-color: #66C6DA" href="'
    )
    # Find all occurences of OUT_PATH and replace it so relative media links work
    out3 = out2.replace(OUT_PATH, "")
    open(outfile_path, "w").write(out3)


def download_from_dropbox():
    if not os.path.exists(OUT_PATH):
        print(f"Creating path [green]{OUT_PATH}[/green]")
        os.mkdir(OUT_PATH)
    if not os.path.exists(EXTRACT_PATH):
        print(f"Creating path [green]{EXTRACT_PATH}[/green]")
        os.mkdir(EXTRACT_PATH)

    print("Downloading zipfile from [blue]Dropbox[/blue]", end=" ")
    # r = requests.get(DROPBOX_URL)
    # if r.status_code == 200:
    if True:
        zip_path = os.path.join(EXTRACT_PATH, "tmp.zip")
        # open(zip_path, "wb").write(r.content)
        with zipfile.ZipFile(zip_path) as Z:
            for zf in track(Z.infolist()):
                if not zf.filename.startswith("__MACOSX"):
                    Z.extract(zf, EXTRACT_PATH)


def collect_paths_todo():
    filepaths = []
    for dirpath, dirnames, filenames in os.walk(EXTRACT_PATH):
        if dirpath.lower().find("werkmap") > 0:
            continue
        for filename in filenames:
            newfilename = filename.replace(" ", "_").lower()
            # We are going to make all filenames lowercase here and remove spaces!
            newfilepath = os.path.join(dirpath, newfilename)
            os.rename(os.path.join(dirpath, filename), newfilepath)
            filepaths.append(newfilepath)
    return filepaths


def fixes(elem, doc):
    if isinstance(elem, panflute.Image):
        elem.url = elem.url.replace(OUT_PATH, "")


def hid(somestring):
    n = zlib.crc32(somestring)
    return "%s" % ("00000000%x" % (n & 0xFFFFFFFF))[-8:]


def convert_docx(input_path):
    if not input_path.lower().endswith(".docx"):
        return
    outfile_media = input_path.replace(EXTRACT_PATH, OUT_PATH).replace(".docx", "")
    data = pypandoc.convert_file(
        input_path,
        "json",
        extra_args=[f"--extract-media={outfile_media}"],
    )
    doc = panflute.load(io.StringIO(data))
    # fix all the href in images
    newdoc = panflute.run_filter(fixes, doc=doc)

    tags = []
    newd = []
    _, docfilename = os.path.split(input_path)
    tmp = {
        "filepath": input_path,
        "doc": newdoc,
        "filename": docfilename.lower().replace(".docx", ""),
    }
    for p in newdoc.content:
        s = panflute.stringify(p).strip()
        if s.lower().startswith("tags:"):
            tmp["tags"] = [
                ss.strip() for ss in s.split(",") if not ss.lower().startswith("tags")
            ]
        else:
            newd.append(p)

    tmp["title"] = panflute.stringify(newd[0]).strip()
    if len(newd) > 1:
        newd = newd[1:]

    tmp["html"] = Markup(
        panflute.convert_text(newd, input_format="panflute", output_format="html")
    )
    tmp["filepath"] = input_path
    return tmp


def to_html(obj):
    if not obj:
        return
    try:
        template = env.get_template(f"post.html")
    except jinja2.exceptions.TemplateNotFound:
        print("Oops, where is your template?")
        sys.exit(1)

    out = template.render(obj)
    filename = obj.get("filename")
    outfile_path = os.path.join(OUT_PATH, filename)
    if not os.path.exists(outfile_path):
        os.mkdir(outfile_path)

    open(outfile_path + "/index.html", "w").write(out)


def main():
    # download_from_dropbox()
    data = []
    tags = {}

    for f in track(collect_paths_todo()):
        try:
            obj = convert_docx(f)
            if not obj:
                continue
            to_html(obj)
            data.append(obj)
            for tag in obj.get("tags", []):
                tags.setdefault(tag, []).append(obj)
        except RuntimeError:
            print(f"Problem with {f}")

    # sort tags by usage
    tags_by_count = reversed(
        sorted([(t, tt) for t, tt in tags.items()], key=lambda x: len(x[1]))
    )

    outfile_path = os.path.join(OUT_PATH, "index.html")
    template = env.get_template("index.html")
    out = template.render({"objs": data, "tags": tags_by_count})
    open(outfile_path, "w").write(out)

    for tag, objs in tags.items():        
        outfile_path = os.path.join(OUT_PATH, f"{tag}.html")
        out = template.render({"objs": objs, "tags": tags_by_count, "tag":tag})
        open(outfile_path, "w").write(out)
    return data


if __name__ == "__main__":
    data = main()

"""
archeologie R 0 / G 78 / B 146
boekgeschiedenis R 233 / G 131 / B 0
boekwetenschap R 0 / G 78 / B 146
cultuurgeschiedenis R 0 / G 78 / B 146
geschiedenis R 233 / G 131 / B 0
kunstgeschiedenis R 0 / G 78 / B 146
virtueel R 0 / G 78 / B 146
"""
