import openai
import base64
import mimetypes
from openai import OpenAI

def perform_ocr_with_gpt4_vision(image_path: str, api_key: str, model: str) -> str:
    """
    Uses the new OpenAI client interface to perform OCR with GPT‑4 Vision.
    
    Parameters:
      image_path (str): Path to the image file.
      api_key (str): Your OpenAI API key.
      model (str): The model to use (e.g. "gpt-4o", "gpt-4o-mini", "gpt-4-turbo").
      
    Returns:
      str: The recognized text, or an empty string if an error occurred.
    """
    # Initialize the OpenAI client with the provided API key.
    client = OpenAI(api_key=api_key)
    
    # Determine MIME type and encode image file as base64 data URL.
    mime_type, _ = mimetypes.guess_type(image_path)
    if mime_type is None:
        mime_type = "application/octet-stream"
    with open(image_path, "rb") as img_file:
        b64_data = base64.b64encode(img_file.read()).decode("utf-8")
    data_url = f"data:{mime_type};base64,{b64_data}"
    
    # Prepare messages for the chat completion.
    messages = [
        {"role": "system", "content": "You are an OCR assistant. Extract all text from the image."},
        {"role": "user", "content": [
            {"type": "image_url", "image_url": {"url": data_url}}
        ]}
    ]
    try:
        response = client.chat.completions.create(model=model, messages=messages)
        recognized_text = response.choices[0].message.content
        return recognized_text
    except Exception as e:
        print("[DEBUG] OpenAI Vision OCR error:", e)
        return ""
