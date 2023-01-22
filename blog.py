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
from docx import Document
import configparser


if len(sys.argv) < 2:
    print("A config file to read should be specified as first argument")
    sys.exit(2)

config = configparser.ConfigParser()
config.read(sys.argv[1])
if "main" not in config:
    print(f"[red]main[/red] section not found in config {sys.argv[1]}")
    sys.exit(1)

main = config["main"]


DROPBOX_URL = main.get("dropbox_url")
TEMPLATE_PATH = main.get("template_path")
OUT_PATH = main.get("out_path")
EXTRACT_PATH = main.get("extract_path")
DOWNLOAD = main.getboolean("download")


print(f"Init template environment [green]{TEMPLATE_PATH}[/green]")
env = Environment(
    loader=FileSystemLoader(TEMPLATE_PATH),
    autoescape=select_autoescape(["html", "xml"]),
)


def go(input_path):
    if (
        input_path.endswith(".jpg")
        or input_path.endswith(".png")
        or input_path.endswith(".pdf")
    ):
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
    zip_path = os.path.join(EXTRACT_PATH, "tmp.zip")
    if os.path.exists(zip_path):
        print("File exists [blue]tmp.zip[/blue], not downloading again.")
        return zip_path

    if not DOWNLOAD:
        print("Download flag set to [blue]False[/blue], not downloading again.")
        return zip_path

    if not os.path.exists(OUT_PATH):
        print(f"Creating path [green]{OUT_PATH}[/green]")
        os.mkdir(OUT_PATH)
    if not os.path.exists(EXTRACT_PATH):
        print(f"Creating path [green]{EXTRACT_PATH}[/green]")
        os.mkdir(EXTRACT_PATH)

    print("Downloading zipfile from [blue]Dropbox[/blue]", end=" ")
    r = requests.get(DROPBOX_URL)
    if r.status_code == 200:
        print("OK! extracting the [blue]tmp.zip[/blue]")
        open(zip_path, "wb").write(r.content)
        with zipfile.ZipFile(zip_path) as Z:
            for zf in track(Z.infolist()):
                if not zf.filename.startswith("__MACOSX"):
                    Z.extract(zf, EXTRACT_PATH)
    return zip_path


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


def fixes_prep(doc):
    doc.first_image = None


def fixes(elem, doc):
    if isinstance(elem, panflute.Image):
        if doc.first_image is None:
            doc.first_image = elem
        elem.url = elem.url.replace(OUT_PATH, "")
        elem.attributes = {}


def hid(somestring):
    n = zlib.crc32(somestring)
    return "%s" % ("00000000%x" % (n & 0xFFFFFFFF))[-8:]


def convert_docx(input_path):
    if (
        input_path.endswith(".jpg")
        or input_path.endswith(".png")
        or input_path.endswith(".pdf")
    ):
        print(f"Copying {input_path} to {OUT_PATH}")
        shutil.copy(input_path, OUT_PATH)

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
    newdoc = panflute.run_filter(fixes, prepare=fixes_prep, doc=doc)

    tags = []
    newd = []
    _, docfilename = os.path.split(input_path)
    docfilename = docfilename.lower().replace(".docx", "")
    try:
        seq = int(docfilename.split("_")[-1])
    except:
        seq = 0

    d = Document(input_path)
    tmp = {
        "seq": seq,
        "filepath": input_path,
        "doc": newdoc,
        "filename": docfilename,
        "author": d.core_properties.author,
        "title": d.core_properties.title,
        "subject": d.core_properties.subject,
    }

    for p in newdoc.content:
        s = panflute.stringify(p).strip()
        if s.lower().startswith("tags:"):
            tmp["tags"] = [
                ss.strip() for ss in s.split(",") if not ss.lower().startswith("tags")
            ]
        else:
            newd.append(p)

    if not tmp["title"]:
        tmp["title"] = panflute.stringify(newd[0]).strip()
        if len(newd) > 1:
            newd = newd[1:]

    tmp["html"] = Markup(
        panflute.convert_text(newd, input_format="panflute", output_format="html")
    )

    tmp["slug"] = sluggify(tmp.get("title", tmp.get("filename", "_")))

    return tmp


def sluggify(something):
    buf = []
    for c in something.lower():
        if c == " ":
            c = "-"
        if c in "0123456789abcdefghijklmnopqrstuvwxyz-":
            buf.append(c)
    tmp = "".join(buf)
    return tmp.strip("-")


def to_html(obj):
    if not obj:
        return

    filename = obj.get("filename")
    try:
        template = env.get_template(f"{filename}.html")
    except jinja2.exceptions.TemplateNotFound:
        try:
            template = env.get_template(f"post.html")
        except jinja2.exceptions.TemplateNotFound:
            print("Oops, where is your template?")
            sys.exit(1)

    out = template.render(obj)
    outfile_path = os.path.join(OUT_PATH, obj.get("slug") + ".html")
    open(outfile_path, "w").write(out)


def main():
    zip_path = download_from_dropbox()
    data = []
    tags = {}
    authors = {}

    for f in track(collect_paths_todo()):
        try:
            obj = convert_docx(f)
            if not obj:
                continue
            to_html(obj)
            data.append(obj)
            for tag in obj.get("tags", []):
                tags.setdefault(tag, []).append(obj)
            if len(obj.get("author", "")) > 0:
                authors.setdefault(obj["author"], []).append(obj)
        except RuntimeError:
            print(f"Problem with {f}")

    # sort tags by usage
    tags_by_count = list(
        reversed(sorted([(t, tt) for t, tt in tags.items()], key=lambda x: len(x[1])))
    )

    outfile_path = os.path.join(OUT_PATH, "index.html")
    template = env.get_template("index.html")
    out = template.render(
        {
            "objs": list(reversed(sorted(data, key=lambda x: x.get("seq", 0))))[:24],
            "tags": [x for x, y in tags_by_count],
        }
    )
    open(outfile_path, "w").write(out)

    template = env.get_template("posts.html")
    for tag, objs in tags.items():
        outfile_path = os.path.join(OUT_PATH, f"{tag}.html")
        out = template.render(
            {
                "objs": reversed(sorted(objs, key=lambda x: x.get("seq", 0))),
                "tags": [x for x, y in tags_by_count],
                "tag": tag,
            }
        )
        open(outfile_path, "w").write(out)

    for author, objs in authors.items():
        outfile_path = os.path.join(OUT_PATH, f"{author}.html")
        out = template.render(
            {
                "objs": reversed(sorted(objs, key=lambda x: x.get("seq", 0))),
                "tags": [x for x, y in tags_by_count],
                "tag": author,
            }
        )
        open(outfile_path, "w").write(out)

    if os.path.exists(zip_path):
        print(f"Deleting [red]{zip_path}[/red]")
        os.remove(zip_path)
    return data


if __name__ == "__main__":
    data = main()
