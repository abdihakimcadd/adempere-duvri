import json
from config import llm_client, LLM_MODEL

def review_duvri(draft_json, frontend_json, context_block):
    prompt = f"""
    You are a strict Italian ASL (Health Inspector) and Legal Auditor for D.Lgs 81/08.
    Review the following DUVRI draft JSON for legal compliance and safety accuracy.
    
    CHECKLIST:
    1. Do the prevention measures strictly follow the Hierarchy of Controls (Elimination -> Collective -> PPE)?
    2. Did the drafter cite Art. 26, D.Lgs 81/08, and INAIL guidelines?
    3. Did the drafter include the mandatory "Valutazione Ricognitiva" (Art. 26, comma 3-ter)?
    4. Did the drafter mention that safety costs are not subject to discount (Art. 26, comma 5)?
    5. Did the drafter invent any risks NOT present in the user's selected risks? If so, flag them.
    
    CONTEXT (SAFETY RULES):
    {context_block}
    
    USER'S ACTUAL SELECTED RISKS:
    {', '.join(frontend_json['risks'])}
    
    DRAFT JSON TO REVIEW:
    {json.dumps(draft_json, ensure_ascii=False)}
    
    OUTPUT (JSON ONLY):
    If the draft is perfect, return the original draft unchanged.
    If there are errors, fix them and return the corrected JSON.
    {{
      "risk_measures": "corrected string",
      "coordination_procedures": "corrected string",
      "safety_costs": "corrected string",
      "valutazione_ricognitiva": "corrected string",
      "review_notes": "string explaining what was fixed, or 'PASS' if no changes needed"
    }}
    """
    response = llm_client.chat.completions.create(
        model=LLM_MODEL,
        messages=[{"role": "user", "content": prompt}],
        response_format={"type": "json_object"}
    )
    return json.loads(response.choices[0].message.content)