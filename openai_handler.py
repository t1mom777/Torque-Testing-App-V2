import openai
import base64
import mimetypes
import json
from openai import OpenAI

def perform_extraction_from_image(image_path: str, api_key: str, model: str) -> dict:
    """
    Uses the OpenAI API to extract specific torque wrench details from an image.
    Expects the API to return a JSON object with keys:
    manufacturer, model, unit, serial, customer, phone, address.
    """
    client = OpenAI(api_key=api_key)
    mime_type, _ = mimetypes.guess_type(image_path)
    if mime_type is None:
        mime_type = "application/octet-stream"
    with open(image_path, "rb") as img_file:
        b64_data = base64.b64encode(img_file.read()).decode("utf-8")
    data_url = f"data:{mime_type};base64,{b64_data}"
    
    messages = [
        {"role": "system", "content": "You are an assistant that extracts specific fields from an image of a torque wrench label."},
        {"role": "user", "content": (
            "Extract the following information from the image: "
            "Torque Wrench Manufacturer, Torque Wrench Model, Torque Wrench Unit Number, "
            "Torque Wrench Serial Number, Customer/Company Name, Phone Number, and Address. "
            "Return your answer as a JSON object with keys: manufacturer, model, unit, serial, customer, phone, address."
        )},
        {"role": "user", "content": [
            {"type": "image_url", "image_url": {"url": data_url}}
        ]}
    ]
    
    try:
        response = client.chat.completions.create(model=model, messages=messages)
        raw_content = response.choices[0].message.content
        print("[DEBUG] Raw API response:", raw_content)  # Debug print to inspect the raw response
        
        # Remove markdown code block formatting if present.
        if raw_content.startswith("```"):
            lines = raw_content.splitlines()
            # Remove the first line if it contains the backticks and optional language identifier
            if lines and lines[0].strip().startswith("```"):
                lines = lines[1:]
            # Remove the last line if it contains only backticks
            if lines and lines[-1].strip().startswith("```"):
                lines = lines[:-1]
            raw_content = "\n".join(lines).strip()
        
        return json.loads(raw_content)
    except Exception as e:
        print("[DEBUG] OpenAI Extraction error:", e)
        return {}
