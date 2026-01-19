from flask import Flask, render_template, request, jsonify
import json
import re
from docx import Document
from io import BytesIO

app = Flask(__name__)

def load_bible():
    with open('rules.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def apply_rules(text):
    bible = load_bible()
    refined = text
    
    # R-01: Typography
    refined = refined.replace("--", "—")

    # R-03 & R-05: Regex substitutions
    refined = re.sub(r"(\d+)(mg|mL|kg|nm|μm)\b", r"\1 \2", refined)
    refined = re.sub(r"(\d+)\s?dollars", r"$\1", refined, flags=re.I)

    # R-02, R-04, R-06: Dictionary Mapping
    mapping_rules = ["R-02", "R-04", "R-06"]
    for r_id in mapping_rules:
        rule = next(r for r in bible['style_rules'] if r['rule_id'] == r_id)
        for old, new in rule.get('map', {}).items():
            refined = re.sub(re.escape(old), new, refined, flags=re.I)
    
    return refined

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    # Handle File Upload
    if 'file' in request.files:
        file = request.files['file']
        if file.filename.endswith('.docx'):
            doc = Document(BytesIO(file.read()))
            content = "\n".join([p.text for p in doc.paragraphs])
        else:
            content = file.read().decode('utf-8')
    # Handle Typed Text
    else:
        content = request.json.get("text", "")

    refined_text = apply_rules(content)
    return jsonify({"original": content, "refined": refined_text})

if __name__ == '__main__':
    app.run(debug=True)