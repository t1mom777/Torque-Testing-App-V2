import openai

def perform_ocr_with_gpt4_vision(image_path: str, api_key: str, model: str) -> str:
    """
    Uses the new ChatCompletion API with file input to perform OCR with GPT‑4 Vision.
    Ensure you are enrolled in the GPT‑4 Vision beta.
    
    Parameters:
      image_path (str): Path to the image file.
      api_key (str): Your OpenAI API key.
      model (str): The model to use (e.g. "gpt-4-turbo").
      
    Returns:
      str: The recognized text, or an empty string if an error occurred.
    """
    openai.api_key = api_key
    try:
        response = openai.ChatCompletion.create(
            model=model,
            messages=[
                {"role": "system", "content": "You are an OCR assistant. Extract all text from the attached image."}
            ],
            file=open(image_path, "rb")
        )
        recognized_text = response.choices[0].message["content"]
        return recognized_text
    except Exception as e:
        print("[DEBUG] OpenAI Vision OCR error:", e)
        return ""
