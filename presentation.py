import json
import os
import re
from pathlib import Path

from google import genai
from pptx import Presentation


def _extract_json(text):
    fenced = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
    candidate = fenced.group(1) if fenced else text[text.find("{") : text.rfind("}") + 1]
    if not candidate:
        raise ValueError("The model did not return JSON.")
    return json.loads(candidate)


def powerpoint(user_query, slide_count=6, audience="general audience", tone="clear and persuasive"):
    api_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
    if not api_key:
        raise RuntimeError("Set GEMINI_API_KEY in key.env or your environment before generating a presentation.")

    client = genai.Client(api_key=api_key)
    prompt = f"""
You are an expert PowerPoint creator. Create a {slide_count}-slide presentation for {audience}.
Topic or thesis: {user_query}
Tone: {tone}

Return only valid JSON with this exact shape:
{{"title": "presentation title", "slides": [
  {{"title": "slide title", "bullets": ["short bullet", "short bullet"]}}
]}}

Use 3 to 5 concise bullets per slide. Do not use markdown or extra commentary.
"""
    response = client.models.generate_content(
        model="gemini-3.1-flash-lite",
        contents=prompt,
    )
    presentation_info = _extract_json(response.text)
    if not presentation_info.get("slides"):
        raise ValueError("The model returned no slides.")
    return presentation_info


def presentation(presentation_info, ppt_file="report.pptx"):
    prs = Presentation()
    title_layout = prs.slide_layouts[0]
    content_layout = prs.slide_layouts[1]

    slides = presentation_info.get("slides", [])
    for index, slide_data in enumerate(slides):
        layout = title_layout if index == 0 else content_layout
        slide = prs.slides.add_slide(layout)
        slide.shapes.title.text = str(slide_data.get("title", ""))

        if index == 0:
            slide.placeholders[1].text = "\n".join(
                str(item) for item in slide_data.get("bullets", [])
            )
            continue

        text_frame = slide.placeholders[1].text_frame
        text_frame.clear()
        for bullet_index, bullet in enumerate(slide_data.get("bullets", [])):
            paragraph = text_frame.paragraphs[0] if bullet_index == 0 else text_frame.add_paragraph()
            paragraph.text = str(bullet)
            paragraph.level = 0

    output_path = Path(ppt_file)
    prs.save(output_path)
    return output_path
