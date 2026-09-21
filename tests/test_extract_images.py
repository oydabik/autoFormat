import zipfile
import xml.etree.ElementTree as ET

from generator import ReportGenerator


RID_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"


def test_extract_images_order(doc_with_15_images):
    with zipfile.ZipFile(doc_with_15_images, "r") as z:
        # 1. Порядок rId из document.xml
        doc_xml = z.read("word/document.xml").decode("utf-8")
        doc_root = ET.fromstring(doc_xml)
        rids = [
            el.get(RID_NS)
            for el in doc_root.iter()
            if el.get(RID_NS) is not None
        ]
        
        # 2. Mapping rId → media/... из rels
        rels_xml = z.read("word/_rels/document.xml.rels").decode("utf-8")
        rels_root = ET.fromstring(rels_xml)
        mapping = {
            rel.get("Id"): rel.get("Target")
            for rel in rels_root.findall(REL_NS)
            if rel.get("Target") and rel.get("Target").startswith("media/")
        }
        
        # 3. Эталонные bytes
        expected_bytes = [z.read(f"word/{mapping[rid]}") for rid in rids]
    
    # 4. Реальные bytes
    gen = ReportGenerator()
    images = gen._extract_images(doc_with_15_images)
    
    # 5. Сравнение
    assert images == expected_bytes