import json

class BibleParser:
    """
    Parses a human-readable Markdown style guide into a structured JSON config.
    """
    @staticmethod
    def parse_markdown_to_json(text: str):
        rules = []
        current_category = "General"
        current_rule = {}
        
        lines = text.strip().split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            # Category Detection
            if line.startswith("# "):
                current_category = line[2:].strip()
                
            # Rule Start Detection
            elif line.startswith("## "):
                if current_rule:
                    rules.append(current_rule)
                
                parts = line[3:].split(":")
                rule_id = parts[0].strip()
                
                current_rule = {
                    "rule_id": rule_id,
                    "category": current_category,
                    "triggers": [],
                    "instruction": "",
                    "exceptions": [],
                    "test_cases": []
                }
                
            # Parsing Details
            elif line.startswith("- Triggers:"):
                raw = line.split(":", 1)[1].strip()
                current_rule["triggers"] = [t.strip() for t in raw.split(",")]
                
            elif line.startswith("- Instruction:"):
                current_rule["instruction"] = line.split(":", 1)[1].strip()
                
            elif line.startswith("- Exceptions:"):
                raw = line.split(":", 1)[1].strip()
                if raw.lower() != "none":
                    current_rule["exceptions"] = [e.strip() for e in raw.split(",")]
                    
            elif line.startswith("- Test Vector:"):
                parts = line.split(":", 1)[1].split("->")
                if len(parts) == 2:
                    current_rule["test_cases"].append({
                        "input": parts[0].strip().strip('"'),
                        "output": parts[1].strip().strip('"')
                    })

        # Append final rule
        if current_rule:
            rules.append(current_rule)
            
        return rules