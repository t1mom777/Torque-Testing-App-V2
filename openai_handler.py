import openai
import base64
import mimetypes
import json
import re
from openai import OpenAI

def perform_extraction_from_image(image_path: str, api_key: str, model: str) -> dict:
    """
    Uses the OpenAI API to extract specific torque wrench details from an image.
    Returns a dictionary with keys:
      manufacturer, model, unit, serial, customer, phone, address, max_torque, torque_unit
    If any key is missing, it will be an empty string.
    """

    # Initialize the OpenAI client
    client = OpenAI(api_key=api_key)

    # Guess the mime type (e.g., "image/png") for the provided file
    mime_type, _ = mimetypes.guess_type(image_path)
    if mime_type is None:
        mime_type = "application/octet-stream"

    # Read the image as base64
    with open(image_path, "rb") as img_file:
        b64_data = base64.b64encode(img_file.read()).decode("utf-8")
    data_url = f"data:{mime_type};base64,{b64_data}"

    # Build the prompt/messages for the ChatCompletion
    messages = [
        {
            "role": "system",
            "content": (
                "You are an assistant that extracts specific fields from an image of a torque wrench label. "
                "Only output valid JSON. Do not include extra commentary or text outside the JSON. "
                "The JSON must have these keys exactly: manufacturer, model, unit, serial, customer, phone, "
                "address, max_torque, torque_unit."
            )
        },
        {
            "role": "user",
            "content": (
                "Extract the following information from the image: "
                "1) Torque Wrench Manufacturer, 2) Torque Wrench Model, 3) Torque Wrench Unit Number, "
                "4) Torque Wrench Serial Number, 5) Customer/Company Name, 6) Phone Number, 7) Address, "
                "8) The maximum torque value (numerical), 9) The torque unit (e.g. ft-lb or Nm). "
                "Return your answer as a JSON object with keys: manufacturer, model, unit, serial, "
                "customer, phone, address, max_torque, torque_unit."
            )
        },
        {
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": data_url}}
            ]
        }
    ]

    # Call the OpenAI ChatCompletion API
    try:
        response = client.chat.completions.create(model=model, messages=messages)
        raw_content = response.choices[0].message.content
        print("[DEBUG] Raw API response:", raw_content)  # Debug log

        # 1) Try to find a JSON code block (```json ... ```).
        match = re.search(
            r'```(?:json)?\s*(\{.*?\})\s*```',
            raw_content,
            flags=re.DOTALL
        )
        if match:
            json_str = match.group(1).strip()
            try:
                data = json.loads(json_str)
            except json.JSONDecodeError:
                # If the block is invalid JSON, fallback to empty
                data = {}
        else:
            # If no code block found, attempt to parse the entire response as JSON
            fallback_content = raw_content.strip('`')
            try:
                data = json.loads(fallback_content)
            except (json.JSONDecodeError, TypeError):
                data = {}

        # 2) Ensure all desired fields exist, defaulting to empty string if missing
        final_data = {
            "manufacturer": data.get("manufacturer", ""),
            "model": data.get("model", ""),
            "unit": data.get("unit", ""),
            "serial": data.get("serial", ""),
            "customer": data.get("customer", ""),
            "phone": data.get("phone", ""),
            "address": data.get("address", ""),
            "max_torque": data.get("max_torque", ""),
            "torque_unit": data.get("torque_unit", "")
        }

        return final_data

    except Exception as e:
        print("[DEBUG] OpenAI Extraction error:", e)
        return {
            "manufacturer": "",
            "model": "",
            "unit": "",
            "serial": "",
            "customer": "",
            "phone": "",
            "address": "",
            "max_torque": "",
            "torque_unit": ""
        }
