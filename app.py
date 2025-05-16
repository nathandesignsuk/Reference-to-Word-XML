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
        title = entry.get("title") or "Untitled"
        tag = entry.get("ID", title)[:15]
        author = entry.get("author", "Unknown Author")
        url = entry.get("url", "")
        year = entry.get("year", "2025")
        urldate = entry.get("urldate", "2025-01-01").split("-")

        guid = str(uuid.uuid4())
        y_accessed, m_accessed, d_accessed = urldate[0], urldate[1] if len(urldate) > 1 else "01", urldate[2] if len(urldate) > 2 else "01"

        source = etree.SubElement(root, f"{{{NS}}}Source")
        etree.SubElement(source, f"{{{NS}}}Tag").text = tag
        etree.SubElement(source, f"{{{NS}}}SourceType").text = "InternetSite"
        etree.SubElement(source, f"{{{NS}}}Guid").text = f"{{{guid}}}"
        etree.SubElement(source, f"{{{NS}}}Title").text = title
        etree.SubElement(source, f"{{{NS}}}Year").text = year
        etree.SubElement(source, f"{{{NS}}}YearAccessed").text = y_accessed
        etree.SubElement(source, f"{{{NS}}}MonthAccessed").text = m_accessed
        etree.SubElement(source, f"{{{NS}}}DayAccessed").text = d_accessed
        etree.SubElement(source, f"{{{NS}}}URL").text = url

        author_block = etree.SubElement(source, f"{{{NS}}}Author")
        corporate = etree.SubElement(author_block, f"{{{NS}}}Author")
        etree.SubElement(corporate, f"{{{NS}}}Corporate").text = author

    return etree.tostring(root, pretty_print=True, xml_declaration=True, encoding="UTF-8")

def parse_ris(text):
    entries = []
    entry = {}
    for line in text.splitlines():
        if line.startswith('TY  -'):
            entry = {}
        elif line.startswith('ER  -'):
            if not entry.get('title'):
                entry['title'] = 'Untitled'
            entries.append(entry)
        else:
            key = line[:2].strip()
            val = line[6:].strip()
            if key == 'TI':
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
                # Join all child text parts
                title = ''.join(child.text or '' for child in title_elem)
        else:
            title = 'Untitled'

        entry = {
            'title': title.strip() or 'Untitled',
            'author': item.findtext('contributors/authors/author', default='Unknown Author'),
            'year': item.findtext('dates/year', default='2025'),
            'url': item.findtext('urls/related', default=''),
            'urldate': item.findtext('dates/accessDate', default='2025-01-01')
        }
        entries.append(entry)
    return entries

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000)
