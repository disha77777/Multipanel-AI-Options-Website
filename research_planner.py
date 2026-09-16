import json
import os
import re

from google import genai


def _extract_json(text):
    """Extract JSON from text, handling markdown code blocks."""
    fenced = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
    candidate = fenced.group(1) if fenced else text[text.find("{") : text.rfind("}") + 1]
    if not candidate:
        raise ValueError("The model did not return JSON.")
    return json.loads(candidate)


def generate_research_plan(
    thesis_statement,
    field_of_study="General Research",
    research_type="Thesis/Dissertation",
    duration_months=12,
):
    """Generate a comprehensive research project plan using Gemini."""
    api_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
    if not api_key:
        raise RuntimeError("Set GEMINI_API_KEY in .env before generating a research plan.")

    client = genai.Client(api_key=api_key)
    prompt = f"""
You are an expert research advisor. Create a detailed research project plan for a {research_type} in {field_of_study}.
Thesis/Research Question: {thesis_statement}
Expected Duration: {duration_months} months

Return only valid JSON with this exact shape:
{{
  "title": "Research Project Title",
  "thesis_statement": "Clear thesis statement",
  "field": "{field_of_study}",
  "objectives": [
    {{"title": "Objective 1", "description": "detailed description"}},
    {{"title": "Objective 2", "description": "detailed description"}}
  ],
  "literature_review": {{
    "scope": "What to cover in literature review",
    "key_themes": ["theme 1", "theme 2", "theme 3"],
    "sources_to_explore": ["source type 1", "source type 2"]
  }},
  "methodology": {{
    "research_design": "Type of research design (qualitative/quantitative/mixed)",
    "data_collection": ["method 1", "method 2"],
    "analysis_approach": "How data will be analyzed",
    "sample_size_or_scope": "Expected scope or sample details"
  }},
  "timeline": [
    {{"phase": "Phase 1", "description": "What", "duration_weeks": 4}},
    {{"phase": "Phase 2", "description": "What", "duration_weeks": 6}}
  ],
  "resources_needed": [
    {{"resource": "Resource name", "justification": "Why needed"}},
    {{"resource": "Resource name", "justification": "Why needed"}}
  ],
  "expected_outcomes": [
    "Primary outcome 1",
    "Primary outcome 2",
    "Deliverable 1"
  ],
  "potential_challenges": [
    {{"challenge": "Challenge name", "mitigation": "How to address"}}
  ],
  "references_to_start": ["Key author 1", "Key concept 1", "Seminal work 1"]
}}

Be specific, actionable, and thorough. Use 3-5 items in each section.
Return ONLY the JSON, no markdown or extra commentary.
"""
    response = client.models.generate_content(
        model="gemini-3.1-flash-lite",
        contents=prompt,
    )
    plan = _extract_json(response.text)

    # Validate essential fields
    if not plan.get("objectives"):
      raise ValueError("The model returned an incomplete research plan.")

    return plan
