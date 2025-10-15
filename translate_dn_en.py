from transformers import MarianMTModel, MarianTokenizer
import pandas as pd
import torch

model_name = "Helsinki-NLP/opus-mt-da-en"
    
# Load model and tokenizer
tokenizer = MarianTokenizer.from_pretrained(model_name)
model = MarianMTModel.from_pretrained(model_name)
device = "cuda" if torch.cuda.is_available() else "cpu"
model = model.to(device)

def translate_danish_to_english(text: str) -> str:
    """
    Translate text from Danish to English using an open-source model from Hugging Face.
    """
    # Tokenize the input
    inputs = tokenizer(text, return_tensors="pt", padding=True).to(device)
    
    # Generate translation
    translated = model.generate(**inputs)
    
    # Decode and return
    english_text = tokenizer.decode(translated[0], skip_special_tokens=True)
    return english_text

if __name__ == "__main__":
    df = pd.read_csv("powerpoint_data.csv")
    df['fixed body'] = df['fixed body'].apply(translate_danish_to_english)
    df['waiting body'] = df['waiting body'].apply(translate_danish_to_english)
    df['not fixed body'] = df['not fixed body'].apply(translate_danish_to_english)
    df['fixed title'] = df['fixed title'].apply(translate_danish_to_english)
    df['waiting title'] = df['waiting title'].apply(translate_danish_to_english)
    df['not fixed title'] = df['not fixed title'].apply(translate_danish_to_english)
    df.to_csv("powerpoint_data_en.csv", index=False)
    print(df.iloc[0])