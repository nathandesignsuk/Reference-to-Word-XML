# Universal Reference to Word Bibliography XML Converter

import os
import uuid
import tempfile
from flask import Flask, request, send_file, render_template_string
from lxml import etree
import bibtexparser
import xml.etree.ElementTree as ETree

app = Flask(__name__)

HTML_FORM = """
<!DOCTYPE html>
<html>
<head><title>Reference to Word XML</title></head>
<body>
  <h2>Reference File to Word XML Converter</h2>
  <p>Supported formats: .bib (BibTeX), .ris (Zotero/Mendeley export), .xml (Mendeley)</p>
  <form method="post" enctype="multipart/form-data">
    <input type="file" name="file" accept=".bib,.ris,.xml">
    <input type="submit" value="Convert">
  </form>
</body>
</html>
"""

@app.route('/', methods=['GET', 'POST'])
def convert_reference():
    if request.method == 'POST':
        file = request.files['file']
        filename = file.filename
        base_filename = os.path.splitext(filename)[0]

        if filename.endswith('.bib'):
            bib_data = bibtexparser.load(file)
            entries = bib_data.entries
        elif filename.endswith('.ris'):
            entries = parse_ris(file.read().decode('utf-8'))
        elif filename.endswith('.xml'):
            entries = parse_mendeley_xml(file.read())
        else:
            entries = []

        xml_output = create_word_xml(entries)
        tmp_path = tempfile.mktemp(suffix=".xml")
        with open(tmp_path, 'wb') as f:
            f.write(xml_output)
        return send_file(tmp_path, as_attachment=True, download_name=f"{base_filename}.xml")

    return render_template_string(HTML_FORM)

def create_word_xml(entries):
    NS = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
    root = etree.Element(f"{{{NS}}}Sources", SelectedStyle="", nsmap={"b": NS})

    for entry in entries:
        tag = entry.get("ID") or (entry.get("title", "Untitled")[:15])
        source = etree.SubElement(root, f"{{{NS}}}Source")

        def add_field(name, value):
            if value:
                etree.SubElement(source, f"{{{NS}}}{name}").text = value

        add_field("Tag", tag)
        add_field("SourceType", "InternetSite")
        add_field("Guid", f"{{{str(uuid.uuid4())}}}")
        add_field("Title", entry.get("title"))
        add_field("Year", entry.get("year"))
        add_field("Month", entry.get("month"))
        add_field("Day", entry.get("day"))

        if "urldate" in entry:
            urldate = entry["urldate"].split("-")
            if len(urldate) > 0:
                add_field("YearAccessed", urldate[0])
            if len(urldate) > 1:
                add_field("MonthAccessed", urldate[1])
            if len(urldate) > 2:
                add_field("DayAccessed", urldate[2])

        add_field("URL", entry.get("url"))

        if "author" in entry:
            author_block = etree.SubElement(source, f"{{{NS}}}Author")
            author_nested = etree.SubElement(author_block, f"{{{NS}}}Author")
            etree.SubElement(author_nested, f"{{{NS}}}Corporate").text = entry["author"]

    return etree.tostring(root, pretty_print=True, xml_declaration=True, encoding="UTF-8")

def parse_ris(text):
    entries = []
    entry = {}
    for line in text.splitlines():
        if line.startswith('TY  -'):
            entry = {}
        elif line.startswith('ER  -'):
            if entry:
                entries.append(entry)
        else:
            key = line[:2].strip()
            val = line[6:].strip()
            if key == 'T1':
                entry['title'] = val
            elif key == 'AU':
                entry['author'] = entry.get('author', '') + (', ' if entry.get('author') else '') + val
            elif key == 'PY':
                entry['year'] = val.split('/')[0]
            elif key == 'UR':
                entry['url'] = val
            elif key == 'Y2':
                entry['urldate'] = val.replace('/', '-')
    return entries

def parse_mendeley_xml(xml_bytes):
    entries = []
    root = ETree.fromstring(xml_bytes)
    for item in root.findall('.//record'):
        title_elem = item.find('titles/title')
        if title_elem is not None:
            if title_elem.text:
                title = title_elem.text
            else:
                title = ''.join(child.text or '' for child in title_elem)
        else:
            title = None

        entry = {
            'title': title.strip() if title else None,
            'author': item.findtext('contributors/authors/author') or None,
            'year': item.findtext('dates/year') or None,
            'month': item.findtext('dates/month') or None,
            'day': item.findtext('dates/day') or None,
            'url': item.findtext('urls/related') or None,
            'urldate': item.findtext('dates/accessDate') or None
        }
        entries.append(entry)
    return entries

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000)
