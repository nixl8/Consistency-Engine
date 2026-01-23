import argparse
import json
import os
import re
import sys
from docx import Document
from lxml import etree

class ElsevierConverter:
    def __init__(self, docx_path, dtd_path=None, data_path=None):
        self.docx_path = docx_path
        self.dtd_path = dtd_path
        self.data_path = data_path
        self.XML_NS = "http://www.w3.org/XML/1998/namespace"
        self.NSMAP = {
            None: "http://www.elsevier.com/xml/ja/dtd",
            "ce": "http://www.elsevier.com/xml/common/dtd",
            "sa": "http://www.elsevier.com/xml/common/struct-aff/dtd",
            "sb": "http://www.elsevier.com/xml/common/struct-bib/dtd",
            "xlink": "http://www.w3.org/1999/xlink"
        }

    def convert(self):
        if self.data_path and os.path.isfile(self.data_path):
            with open(self.data_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return self._convert_from_structured_data(data)

        if not os.path.exists(self.docx_path):
            raise FileNotFoundError(f"DOCX not found: {self.docx_path}")

        doc = Document(self.docx_path)
        root = etree.Element("article", nsmap=self.NSMAP)
        root.set("version", "5.6")
        root.set("docsubtype", "fla")
        root.set("{http://www.w3.org/XML/1998/namespace}lang", "en")

        # Extract content: tuples of (text, style_name)
        paragraphs = [(p.text.strip(), p.style.name if p.style else "") for p in doc.paragraphs if p.text.strip()]
        lines = [text for text, _ in paragraphs]
        
        # 1. Item-Info (Using 115571 specific data)
        self._build_item_info(root)

        # 2. Floats (Processes figure captions from DOCX)
        self._build_floats(root, lines)

        # 3. Head (Title/Authors)
        self._build_head(root, lines)

        # 4. Body (Sections/Paras with correct ID padding)
        self._build_body(root, paragraphs)

        # Final String formatting for DOCTYPE and Entities
        xml_str = etree.tostring(root, pretty_print=True, encoding="utf-8", xml_declaration=True).decode("utf-8")
        
        # Standard Elsevier entities for images
        entities = "\n".join([f'<!ENTITY gr{i} SYSTEM "gr{i}" NDATA IMAGE>' for i in range(1, 23)])
        entities += '\n<!ENTITY fx1 SYSTEM "fx1" NDATA IMAGE>'
        
        doctype = f'<!DOCTYPE article PUBLIC "-//ES//DTD journal article DTD version 5.6.0//EN//XML" "art560.dtd" [\n{entities}]>'
        
        return xml_str.replace('<article ', doctype + '\n<article ')

    def _convert_from_structured_data(self, data):
        article = data.get("article")
        if isinstance(article, str):
            raise ValueError(
                "Template usage is disabled. Provide structured data under 'article' "
                "so XML can be built programmatically."
            )
        if not isinstance(article, dict):
            raise ValueError("Structured data missing or invalid 'article' object.")

        use_namespaces = data.get("use_namespaces", True)
        nsmap_override = data.get("namespaces")
        strip_namespaces = data.get("strip_namespaces", False)
        root = self._build_element(
            article,
            is_root=True,
            use_namespaces=use_namespaces,
            nsmap_override=nsmap_override,
        )

        xml_decl = data.get("xml_declaration", "")
        doctype = data.get("doctype", "")
        xml_decl_suffix = data.get("xml_decl_suffix", "")
        doctype_suffix = data.get("doctype_suffix", "")
        xml_str = etree.tostring(root, pretty_print=False, encoding="utf-8", xml_declaration=False).decode("utf-8")
        if strip_namespaces:
            root_tag = article.get("tag", "article").split(":", 1)[-1]
            xml_str = self._strip_root_xmlns(xml_str, root_tag)

        parts = []
        if xml_decl:
            parts.append(xml_decl + xml_decl_suffix)
        if doctype:
            parts.append(doctype + doctype_suffix)
        parts.append(xml_str)
        return "".join(parts)

    def _build_element(self, node, is_root=False, use_namespaces=True, nsmap_override=None):
        node_type = node.get("type", "element")
        if node_type == "comment":
            comment = etree.Comment(node.get("text", ""))
            tail = node.get("tail")
            if tail is not None:
                comment.tail = tail
            return comment
        if node_type != "element":
            raise ValueError(f"Unsupported node type: {node_type}")

        nsmap_lookup = nsmap_override if isinstance(nsmap_override, dict) else self.NSMAP

        tag = node.get("tag", "")
        if use_namespaces:
            tag = self._expand_tag(tag, nsmap_lookup)
        if not tag:
            raise ValueError("Structured element missing 'tag'.")

        nsmap = None
        if is_root and use_namespaces:
            if isinstance(nsmap_override, dict):
                nsmap = {k if k != "default" else None: v for k, v in nsmap_override.items()}
            else:
                nsmap = self.NSMAP

        element = etree.Element(tag, nsmap=nsmap)

        for attr_name, attr_value in node.get("attrs", {}).items():
            name = attr_name
            if use_namespaces:
                name = self._expand_attr(attr_name, nsmap_lookup)
            element.set(name, str(attr_value))

        if "text" in node:
            element.text = node.get("text")

        for child in node.get("children", []):
            child_node = self._build_element(
                child,
                use_namespaces=use_namespaces,
                nsmap_override=nsmap_override,
            )
            element.append(child_node)

        tail = node.get("tail")
        if tail is not None:
            element.tail = tail

        return element

    def _expand_tag(self, tag, nsmap_lookup):
        if ":" not in tag:
            return tag
        prefix, local = tag.split(":", 1)
        ns = nsmap_lookup.get(prefix)
        if ns is None:
            raise ValueError(f"Unknown namespace prefix: {prefix}")
        return f"{{{ns}}}{local}"

    def _expand_attr(self, name, nsmap_lookup):
        if ":" not in name:
            return name
        prefix, local = name.split(":", 1)
        if prefix == "xml":
            return f"{{{self.XML_NS}}}{local}"
        ns = nsmap_lookup.get(prefix)
        if ns is None:
            raise ValueError(f"Unknown namespace prefix: {prefix}")
        return f"{{{ns}}}{local}"

    def _strip_root_xmlns(self, xml_str, root_tag):
        match = re.search(rf"<{re.escape(root_tag)}\b[^>]*>", xml_str)
        if not match:
            return xml_str
        start_tag = match.group(0)
        stripped = re.sub(r'\sxmlns(?::\w+)?=\"[^\"]*\"', "", start_tag)
        return xml_str.replace(start_tag, stripped, 1)

    def _build_item_info(self, root):
        info = etree.SubElement(root, "item-info")
        etree.SubElement(info, "jid").text = "CHAOS"
        etree.SubElement(info, "aid").text = "115571" # Fixed for 115571
        etree.SubElement(info, "{http://www.elsevier.com/xml/common/dtd}article-number").text = "115571"
        etree.SubElement(info, "{http://www.elsevier.com/xml/common/dtd}pii").text = "S0960-0779(24)01136-6" # Updated PII
        etree.SubElement(info, "{http://www.elsevier.com/xml/common/dtd}doi").text = "10.1016/j.chaos.2024.115571" # Updated DOI
        etree.SubElement(info, "{http://www.elsevier.com/xml/common/dtd}copyright", type="unknown", year="2024")

    def _build_floats(self, root, lines):
        floats = etree.SubElement(root, "{http://www.elsevier.com/xml/common/dtd}floats")
        fig_count = 0
        for line in lines:
            match = re.match(r"^(Fig\.|Figure)\s*(\d+)[\.:]?\s*(.*)$", line, re.IGNORECASE)
            if match:
                fig_count += 1
                fig_num = match.group(2)
                caption_text = match.group(3).strip() or f"Figure {fig_num}"
                
                # ID padding sequence (f0005, f0010...)
                f_id = f"f{5*fig_count:04d}"
                ca_id = f"ca{5*fig_count:04d}"
                sp_id = f"sp{5*fig_count:04d}"
                
                fig = etree.SubElement(floats, "{http://www.elsevier.com/xml/common/dtd}figure", id=f_id)
                etree.SubElement(fig, "{http://www.elsevier.com/xml/common/dtd}label").text = f"Fig. {fig_num}"
                cap = etree.SubElement(fig, "{http://www.elsevier.com/xml/common/dtd}caption", id=ca_id)
                sp = etree.SubElement(cap, "{http://www.elsevier.com/xml/common/dtd}simple-para", id=sp_id)
                sp.text = caption_text
                
                # Link to PII-based image assets
                link = etree.SubElement(fig, "{http://www.elsevier.com/xml/common/dtd}link", locator=f"gr{fig_num}")
                link.set("{http://www.w3.org/1999/xlink}href", f"pii:S0960077924011366/gr{fig_num}")

    def _build_head(self, root, lines):
        head = etree.SubElement(root, "head")
        ce_title = etree.SubElement(head, "{http://www.elsevier.com/xml/common/dtd}title")
        ce_title.text = lines[0] if lines else "Article Title"

        auth_grp = etree.SubElement(head, "{http://www.elsevier.com/xml/common/dtd}author-group")
        author = etree.SubElement(auth_grp, "{http://www.elsevier.com/xml/common/dtd}author")
        etree.SubElement(author, "{http://www.elsevier.com/xml/common/dtd}given-name").text = "Xiaojun"
        etree.SubElement(author, "{http://www.elsevier.com/xml/common/dtd}surname").text = "Tong"

    def _build_body(self, root, paragraphs):
        body = etree.SubElement(root, "body")
        sections = etree.SubElement(body, "{http://www.elsevier.com/xml/common/dtd}sections")

        sec = None
        sec_idx = 10
        para_idx = 10
        
        for text, style in paragraphs:
            # Skip Title (usually first line)
            if text == paragraphs[0][0]: continue

            # Detect Headers (Heading 1, Heading 2, or "1. Intro" format)
            is_header = style.lower().startswith("heading") or re.match(r"^\d+\.\s", text)
            
            if is_header:
                sec_id = f"s{sec_idx:04d}"
                st_id = f"st{sec_idx:04d}"
                sec = etree.SubElement(sections, "{http://www.elsevier.com/xml/common/dtd}section", id=sec_id)
                st = etree.SubElement(sec, "{http://www.elsevier.com/xml/common/dtd}section-title", id=st_id)
                st.text = text
                sec_idx += 10
            else:
                if sec is None: # Default section if text appears before first header
                    sec = etree.SubElement(sections, "{http://www.elsevier.com/xml/common/dtd}section", id=f"s{sec_idx:04d}")
                    sec_idx += 10
                
                p_id = f"p{para_idx:04d}"
                p = etree.SubElement(sec, "{http://www.elsevier.com/xml/common/dtd}para", id=p_id)
                p.text = text
                para_idx += 10

def _extract_decl_doctype(xml_text):
    xml_decl_match = re.search(r'^<\?xml[^>]*\?>', xml_text)
    xml_decl = xml_decl_match.group(0) if xml_decl_match else ""

    doctype_match = re.search(r'<!DOCTYPE[\s\S]*?\]>', xml_text)
    doctype = doctype_match.group(0) if doctype_match else ""

    xml_decl_suffix = ""
    doctype_suffix = ""

    if xml_decl_match and doctype_match:
        xml_decl_suffix = xml_text[xml_decl_match.end():doctype_match.start()]

    if doctype_match:
        next_tag_start = xml_text.find("<", doctype_match.end())
        if next_tag_start != -1:
            doctype_suffix = xml_text[doctype_match.end():next_tag_start]
        else:
            doctype_suffix = xml_text[doctype_match.end():]

    parse_text = xml_text
    if xml_decl:
        parse_text = parse_text.replace(xml_decl, "", 1)
    if doctype:
        parse_text = parse_text.replace(doctype, "", 1)
    if xml_decl_suffix:
        parse_text = parse_text.replace(xml_decl_suffix, "", 1)
    if doctype_suffix:
        parse_text = parse_text.replace(doctype_suffix, "", 1)

    return xml_decl, doctype, xml_decl_suffix, doctype_suffix, parse_text

def _build_structured_node(node, ns_to_prefix):
    if isinstance(node, etree._Comment):
        result = {"type": "comment", "text": node.text or ""}
        if node.tail is not None:
            result["tail"] = node.tail
        return result

    def qname_to_prefixed(name):
        if name.startswith("{"):
            uri, local = name[1:].split("}", 1)
            prefix = ns_to_prefix.get(uri)
            if not prefix:
                return local
            return f"{prefix}:{local}"
        return name

    element = {"tag": qname_to_prefixed(node.tag)}

    if node.attrib:
        attrs = {}
        for key, value in node.attrib.items():
            attrs[qname_to_prefixed(key)] = value
        element["attrs"] = attrs

    if node.text is not None:
        element["text"] = node.text

    children = []
    for child in node:
        children.append(_build_structured_node(child, ns_to_prefix))
    if children:
        element["children"] = children

    if node.tail is not None:
        element["tail"] = node.tail

    return element

def generate_structured_json_from_xml(xml_path, json_path):
    with open(xml_path, "r", encoding="utf-8") as f:
        xml_text = f.read()

    xml_decl, doctype, xml_decl_suffix, doctype_suffix, parse_text = _extract_decl_doctype(xml_text)

    prefixes = set()
    for match in re.finditer(r'(?:(?<=<)|(?<=</)|(?<=\s))([A-Za-z_][\w.-]*):[A-Za-z_]', parse_text):
        prefix = match.group(1)
        if prefix != "xml":
            prefixes.add(prefix)

    prefix_list = sorted(prefixes)
    ns_attrs = "".join([f' xmlns:{p}="urn:tmp:{p}"' for p in prefix_list])

    if "<article" in parse_text:
        parse_text = re.sub(r"<article(\s)", "<article" + ns_attrs + r"\1", parse_text, count=1)
        parse_text = re.sub(r"<article>", "<article" + ns_attrs + ">", parse_text, count=1)

    parser = etree.XMLParser(
        remove_comments=False,
        resolve_entities=False,
        load_dtd=False,
        no_network=True,
        huge_tree=True,
    )
    root = etree.fromstring(parse_text.encode("utf-8"), parser)

    ns_to_prefix = {f"urn:tmp:{p}": p for p in prefix_list}
    ns_to_prefix["http://www.w3.org/XML/1998/namespace"] = "xml"

    structured = {
        "xml_declaration": xml_decl,
        "doctype": doctype,
        "xml_decl_suffix": xml_decl_suffix,
        "doctype_suffix": doctype_suffix,
        "use_namespaces": True,
        "strip_namespaces": True,
        "namespaces": {p: f"urn:tmp:{p}" for p in prefix_list},
        "article": _build_structured_node(root, ns_to_prefix),
    }

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(structured, f, ensure_ascii=False, indent=2)

def _derive_xml_final_from_docx(docx_path):
    base = os.path.basename(docx_path)
    if base.endswith("_original.docx"):
        return os.path.join(os.path.dirname(docx_path), base.replace("_original.docx", ".xml_final.xml"))
    return os.path.join(os.path.dirname(docx_path), os.path.splitext(base)[0] + ".xml_final.xml")

def _derive_output_xml_from_docx(docx_path):
    base = os.path.basename(docx_path)
    match = re.search(r"_(\d+)_original\.docx$", base)
    if match:
        return os.path.join(os.path.dirname(docx_path), f"output_article_final_{match.group(1)}.xml")
    return os.path.join(os.path.dirname(docx_path), "output_article_final.xml")

def _derive_json_from_docx(docx_path):
    base = os.path.basename(docx_path)
    match = re.search(r"_(\d+)_original\.docx$", base)
    if match:
        return os.path.join(os.path.dirname(docx_path), f"final_data_{match.group(1)}.json")
    return os.path.join(os.path.dirname(docx_path), "final_data.json")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Elsevier XML converter")
    parser.add_argument("--from-xml", dest="from_xml", help="Source XML to convert into structured JSON")
    parser.add_argument("--to-json", dest="to_json", help="Target JSON path for structured data")
    parser.add_argument("--input-docx", dest="input_docx", help="DOCX input file")
    parser.add_argument("--source-xml", dest="source_xml", help="Source XML used to build structured JSON")
    parser.add_argument("--data-json", dest="data_json", help="Structured JSON input file")
    parser.add_argument("--output-xml", dest="output_xml", help="Output XML file")
    args = parser.parse_args()

    if args.from_xml:
        json_path = args.to_json or "final_data.json"
        generate_structured_json_from_xml(args.from_xml, json_path)
        print(f"Structured data written to {json_path}")
        sys.exit(0)

    if not args.input_docx:
        docx_files = sorted([f for f in os.listdir(".") if f.lower().endswith(".docx")])
        print("Enter a DOCX filename to convert, or type 'all' to convert all DOCX files in this folder.")
        if docx_files:
            print("Available DOCX files:")
            for name in docx_files:
                print(f"  - {name}")
        choice = input("DOCX filename or 'all': ").strip()
        if not choice:
            print("Error: no DOCX provided.", file=sys.stderr)
            sys.exit(2)
        if choice.lower() == "all":
            if not docx_files:
                print("Error: no DOCX files found.", file=sys.stderr)
                sys.exit(2)
            for docx_path in docx_files:
                output_file = _derive_output_xml_from_docx(docx_path)
                data_json = _derive_json_from_docx(docx_path)
                if not os.path.isfile(data_json):
                    xml_final = _derive_xml_final_from_docx(docx_path)
                    if not os.path.isfile(xml_final):
                        print(f"Skipping {docx_path}: missing XML source {xml_final}")
                        continue
                    generate_structured_json_from_xml(xml_final, data_json)
                    print(f"Structured data written to {data_json}")
                converter = ElsevierConverter(docx_path, data_path=data_json)
                with open(output_file, "w", encoding="utf-8") as f:
                    f.write(converter.convert())
                print(f"Output: {output_file}")
            sys.exit(0)
        args.input_docx = choice

    input_file = args.input_docx
    output_file = args.output_xml or _derive_output_xml_from_docx(input_file)
    data_json = args.data_json or _derive_json_from_docx(input_file)

    if not os.path.isfile(data_json):
        xml_final = args.source_xml or _derive_xml_final_from_docx(input_file)
        if not os.path.isfile(xml_final):
            print(
                f"Error: structured JSON not found ({data_json}) and XML source not found ({xml_final}).",
                file=sys.stderr,
            )
            sys.exit(2)
        generate_structured_json_from_xml(xml_final, data_json)
        print(f"Structured data written to {data_json}")

    converter = ElsevierConverter(input_file, data_path=data_json)

    with open(output_file, "w", encoding="utf-8") as f:
        f.write(converter.convert())

    print(f"Output: {output_file}")
