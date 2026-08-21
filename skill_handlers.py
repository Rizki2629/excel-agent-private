"""
Skill handlers for Excel Agent.
Detects user intent and provides skill-specific context and examples.
Based on MCP Server Excel best practices.
"""
import json
import os

SKILLS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'skills.json')

_skills_cache = None


def load_skills():
    """Load skills from JSON file."""
    global _skills_cache
    if _skills_cache is None:
        try:
            with open(SKILLS_FILE, 'r', encoding='utf-8') as f:
                _skills_cache = json.load(f)
        except FileNotFoundError:
            _skills_cache = {"skills": []}
    return _skills_cache


def detect_skill(user_message):
    """Detect which skill matches the user's request."""
    data = load_skills()
    msg = user_message.lower()
    
    best_match = None
    best_score = 0
    
    for skill in data.get("skills", []):
        score = 0
        for trigger in skill.get("triggers", []):
            if trigger.lower() in msg:
                # Longer trigger matches are more specific
                score += len(trigger.split())
        
        if score > best_score:
            best_score = score
            best_match = skill
    
    return best_match


def get_skill_context(skill):
    """Get context string for a detected skill."""
    if not skill:
        return ""
    
    ctx = f"\n\n### SKILL: {skill['name'].upper()} ###\n"
    ctx += f"Deskripsi: {skill['description']}\n\n"
    
    # Workflow
    if skill.get("workflow"):
        ctx += "Workflow:\n"
        for step in skill["workflow"]:
            ctx += f"  {step}\n"
        ctx += "\n"
    
    # Critical rules
    if skill.get("criticalRules"):
        ctx += "Aturan Kritis:\n"
        for rule in skill["criticalRules"]:
            ctx += f"  - {rule}\n"
        ctx += "\n"
    
    # Format codes (if any)
    if skill.get("formatCodes"):
        ctx += "Format Codes:\n"
        for name, code in skill["formatCodes"].items():
            ctx += f"  {name}: {code}\n"
        ctx += "\n"
    
    # Chart types (if any)
    if skill.get("chartTypes"):
        ctx += "Chart Types: " + ", ".join(skill["chartTypes"][:10]) + "\n\n"
    
    # Examples
    if skill.get("examples"):
        ctx += "Contoh Kode:\n"
        for ex in skill["examples"][:2]:  # Max 2 examples
            ctx += f"Request: {ex['request']}\n"
            ctx += f"```javascript\n{ex['code']}\n```\n\n"
    
    # Office APIs used
    if skill.get("officeApis"):
        ctx += f"Office.js APIs: {', '.join(skill['officeApis'])}\n"
    
    return ctx


def get_all_skills_summary():
    """Get summary of all available skills."""
    data = load_skills()
    lines = []
    for skill in data.get("skills", []):
        triggers = ", ".join(skill.get("triggers", [])[:3])
        lines.append(f"- {skill['name']}: {skill['description'][:60]}... (triggers: {triggers})")
    return "\n".join(lines)


if __name__ == "__main__":
    # Test
    test_messages = [
        "Format kolom C sebagai currency",
        "Buat chart penjualan bulanan",
        "Urutkan berdasarkan kolom B",
        "Buat pivot table",
        "Atur lebar kolom A jadi 10",
        "Highlight sel yang lebih dari 80"
    ]
    
    for msg in test_messages:
        skill = detect_skill(msg)
        if skill:
            print(f"Message: {msg}")
            print(f"  Detected skill: {skill['name']}")
            print()
