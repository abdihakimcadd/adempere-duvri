import json
from config import llm_client, LLM_MODEL

def draft_duvri(frontend_json, context_block):
    prompt = f"""
    SYSTEM: You are an Italian HSE Manager. Draft prevention measures for a DUVRI based on D.Lgs 81/08.
    
    CRITICAL RULES:
    1. Write professional, legally binding Italian text.
    2. Cite "D.Lgs 81/08" where appropriate.
    3. Use the EXACT risk scores provided. Do not calculate scores.
    4. Follow the Hierarchy of Controls (Elimination -> Collective Protection -> PPE).
    5. Write all risk names in perfect Italian.
    6. CONTENT RICHNESS: Generate 3 to 4 detailed sentences per section. Describe specific measures for each risk individually.
    7. OUTPUT FORMAT: Output MUST be plain text only. Do NOT use Markdown formatting (no #, *, |, -) or HTML tags. Output pure Italian sentences.
    
    CONTEXT (SAFETY RULES):
    {context_block}
    
    USER FACTS:
    - Host Company: {frontend_json['host_company']}
    - Contractor: {frontend_json['contractor_name']}
    - Work Type: {frontend_json['work_type']}
    - Selected Risks (Italian): {', '.join(frontend_json.get('risks_it', frontend_json['risks']))}
    - Risk Scores: {frontend_json['risk_scores']}
    
    OUTPUT FORMAT (JSON ONLY):
    {{
      "risk_measures": "3-4 sentences detailing prevention measures and PPE for the risks, following Hierarchy of Controls.",
      "coordination_procedures": "3-4 sentences detailing daily briefings, emergency procedures, and information exchange (Art. 26 comma 3).",
      "safety_costs": "3-4 sentences stating costs are identified per Art. 26 comma 5, are not subject to discount, and distinguishing special vs ordinary costs.",
      "valutazione_ricognitiva": "3-4 sentences stating the standard risks for the work type, the methodology used, and specific findings (Art. 26 comma 3-ter)."
    }}
    """
    response = llm_client.chat.completions.create(
        model=LLM_MODEL,
        messages=[{"role": "user", "content": prompt}],
        response_format={"type": "json_object"}
    )
    return json.loads(response.choices[0].message.content)