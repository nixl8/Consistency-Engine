import os
import json
import docx
import re
import urllib.request
import urllib.error
from io import BytesIO
from flask import Flask, render_template, request, jsonify, send_file
from dotenv import load_dotenv
from openai import OpenAI

# Load environment variables
load_dotenv()

app = Flask(__name__, template_folder='.')

# Ensure API configuration is present
client = None
try:
    client = OpenAI(
        api_key=os.getenv("LLMFOUNDRY_TOKEN"),
        base_url=os.getenv("BASE_URL"),
        default_headers={"X-Project-Id": os.getenv("PROJECT_ID")}
    )
except Exception as e:
    print(f"CRITICAL: Failed to initialize OpenAI client: {e}")

def build_foundry_api_key():
    token = os.getenv("LLMFOUNDRY_TOKEN")
    project_id = os.getenv("PROJECT_ID")
    if not token:
        return None
    if project_id:
        return f"{token}:{project_id}"
    return token

foundry_openai_client = None
foundry_groq_client = None
foundry_key = build_foundry_api_key()
if foundry_key:
    foundry_openai_client = OpenAI(
        api_key=foundry_key,
        base_url="https://llmfoundry.straive.com/openai/v1/"
    )
    foundry_groq_client = OpenAI(
        api_key=foundry_key,
        base_url="https://llmfoundry.straive.com/groq/openai/v1/"
    )

# Switched to 1.5-flash as it is more stable for JSON structured output in this tier
MODEL_NAME = "gemini-2.0-flash" 
ACTIVE_CONFIG = []

def clean_llm_json(raw_str):
    """Removes markdown backticks and cleans JSON strings."""
    return re.sub(r'```json\s*|\s*```', '', raw_str).strip()

def normalize_micro_text(value):
    if not isinstance(value, str):
        return value
    value = value.replace("Î¼", "μ")
    value = value.replace("Âµ", "μ")
    value = value.replace("µ", "μ")
    return value

def normalize_config(rules):
    for rule in rules:
        for key in ("rule_id", "category", "instruction"):
            if key in rule:
                rule[key] = normalize_micro_text(rule[key])
        for key in ("triggers", "exceptions"):
            if key in rule and isinstance(rule[key], list):
                rule[key] = [normalize_micro_text(item) for item in rule[key]]
        if "test_cases" in rule and isinstance(rule["test_cases"], list):
            for test_case in rule["test_cases"]:
                if "input" in test_case:
                    test_case["input"] = normalize_micro_text(test_case["input"])
                if "output" in test_case:
                    test_case["output"] = normalize_micro_text(test_case["output"])
    return rules

def normalize_output(payload):
    if not isinstance(payload, dict):
        return payload
    if "corrected_text" in payload:
        payload["corrected_text"] = normalize_micro_text(payload["corrected_text"])
    if "changes" in payload and isinstance(payload["changes"], list):
        for change in payload["changes"]:
            if "original" in change:
                change["original"] = normalize_micro_text(change["original"])
            if "new" in change:
                change["new"] = normalize_micro_text(change["new"])
    return payload

def get_client(provider: str):
    if provider == "llmfoundry":
        return client
    if provider == "openai":
        return foundry_openai_client
    if provider == "groq":
        return foundry_groq_client
    return None

def provider_status():
    return {
        "llmfoundry": client is not None,
        "openai": foundry_openai_client is not None,
        "groq": foundry_groq_client is not None,
        "vertexai-anthropic": foundry_key is not None
    }

def build_system_prompt():
    rules = normalize_config(ACTIVE_CONFIG)
    return f"""
    You are a Deterministic Copyeditor. 
    RULES: {json.dumps(rules, ensure_ascii=False)}
    
    PROTOCOL:
    1. Scan input for every trigger in the JSON rules. 
    2. Pay special attention to apostrophes in years (1990's) and unit symbols (uL).
    3. ONLY modify if text violates the instruction.
    4. Provide the 'corrected_text' with all changes applied.
    
    Return ONLY JSON:
    {{
        "corrected_text": "...",
        "changes": [
            {{ "rule_id": "...", "original": "...", "new": "..." }}
        ]
    }}
    """

def run_completion(provider: str, model: str, user_input: str, system_prompt: str):
    if provider == "vertexai-anthropic":
        return run_vertex_anthropic(model, user_input, system_prompt)

    active_client = get_client(provider)
    if not active_client:
        raise ValueError(f"Provider '{provider}' is not configured. Missing API key or unsupported provider.")

    completion = active_client.chat.completions.create(
        model=model,
        messages=[{"role": "system", "content": system_prompt},
                  {"role": "user", "content": user_input}],
        response_format={"type": "json_object"},
        temperature=0
    )
    return clean_llm_json(completion.choices[0].message.content)

def post_json(url, payload, api_key):
    data = json.dumps(payload).encode("utf-8")
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json; charset=utf-8"
    }
    request_obj = urllib.request.Request(url, data=data, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(request_obj) as response:
            return response.read().decode("utf-8", errors="ignore")
    except urllib.error.HTTPError as err:
        body = err.read().decode("utf-8", errors="ignore")
        raise ValueError(f"Vertex API error {err.code}: {body}")

def safe_json_loads(raw_text):
    if not isinstance(raw_text, str):
        return json.loads(raw_text)
    # Strip ASCII control chars that can appear unescaped in model responses.
    cleaned = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", raw_text)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        return json.loads(cleaned, strict=False)

def run_vertex_anthropic(model: str, user_input: str, system_prompt: str):
    if not foundry_key:
        raise ValueError("Provider 'vertexai-anthropic' is not configured. Missing LLMFOUNDRY_TOKEN/PROJECT_ID.")

    url = f"https://llmfoundry.straive.com/vertexai/anthropic/models/{model}:rawPredict"
    payload = {
        "anthropic_version": "vertex-2023-10-16",
        "max_tokens": 2048,
        "temperature": 0,
        "system": system_prompt,
        "messages": [
            {"role": "user", "content": [{"type": "text", "text": user_input}]}
        ]
    }

    body = post_json(url, payload, foundry_key)
    data = safe_json_loads(body)

    content = ""
    if isinstance(data, dict):
        if isinstance(data.get("content"), list):
            content = "".join(
                part.get("text", "")
                for part in data["content"]
                if isinstance(part, dict)
            )
        else:
            content = data.get("output", "") or data.get("text", "")

    if not content:
        raise ValueError("Vertex Anthropic returned empty response.")

    return clean_llm_json(content)

def extract_text(file):
    """Extracts text from .docx or .txt only."""
    filename = file.filename.lower()
    try:
        if filename.endswith('.docx'):
            doc = docx.Document(file)
            return "\n".join([para.text for para in doc.paragraphs])
        elif filename.endswith('.txt'):
            return file.read().decode('utf-8', errors='ignore')
    except Exception as e:
        print(f"File Extraction Error: {e}")
    return None

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/upload_and_build', methods=['POST'])
def upload_and_build():
    global ACTIVE_CONFIG
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    
    file = request.files['file']
    raw_content = extract_text(file)
    
    if raw_content is None:
        return jsonify({"error": "Only .txt and .docx files are allowed"}), 400

    # ENHANCED BUILDER PROMPT: Explicitly instructs AI to find punctuation/symbol rules
    builder_prompt = f"""
    You are a Universal Style Architect. Analyze the document and extract technical consistency rules.
    CRITICAL: Look for patterns involving:
    1. Spelling: British vs American (e.g., modelled vs modeled).
    2. Decades: Use of apostrophes (e.g., 1990's vs 1990s).
    3. Units: Symbols and spacing (e.g., uL vs μL, 8 % vs 8%).
    4. Punctuation: Oxford commas in lists.
    
    Return ONLY a JSON object:
    {{
      "rules": [
        {{
          "rule_id": "unique_id",
          "category": "category_name",
          "triggers": ["incorrect_example"],
          "instruction": "correction_instruction"
        }}
      ]
    }}
    """

    try:
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[{"role": "system", "content": builder_prompt},
                      {"role": "user", "content": raw_content[:15000]}],
            response_format={"type": "json_object"}
        )
        data = safe_json_loads(clean_llm_json(response.choices[0].message.content))
        ACTIVE_CONFIG = normalize_config(data.get("rules", []))
        return jsonify({"status": "success", "config": ACTIVE_CONFIG})
    except Exception as e:
        print(f"Build Error: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/process_text', methods=['POST'])
def process_text():
    if not ACTIVE_CONFIG:
        return jsonify({"error": "No rules loaded"}), 400

    user_input = request.json.get('text', '')
    
    # ENHANCED SYSTEM PROMPT: Forces strict adherence to the punctuation and symbol triggers
    system_prompt = build_system_prompt()

    try:
        output = run_completion("llmfoundry", MODEL_NAME, user_input, system_prompt)
        payload = safe_json_loads(output)
        return jsonify(normalize_output(payload))
    except Exception as e:
        print(f"Process Error: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/process_text_compare', methods=['POST'])
def process_text_compare():
    if not ACTIVE_CONFIG:
        return jsonify({"error": "No rules loaded"}), 400

    payload = request.json or {}
    user_input = payload.get('text', '')
    targets = payload.get('targets', [])

    if not user_input:
        return jsonify({"error": "No text provided"}), 400
    if not targets:
        return jsonify({"error": "No model targets provided"}), 400

    system_prompt = build_system_prompt()
    results = {}
    errors = {}

    for target in targets:
        provider = (target.get('provider') or "llmfoundry").strip()
        model = (target.get('model') or MODEL_NAME).strip()
        key = f"{provider}:{model}"

        try:
            output = run_completion(provider, model, user_input, system_prompt)
            results[key] = normalize_output(safe_json_loads(output))
        except Exception as e:
            errors[key] = str(e)

    return jsonify({"results": results, "errors": errors})

@app.route('/providers', methods=['GET'])
def providers():
    return jsonify(provider_status())

if __name__ == '__main__':
    try:
        app.run(debug=True, port=5000)
    except Exception as startup_error:
        print(f"SERVER STARTUP FAILED: {startup_error}")
