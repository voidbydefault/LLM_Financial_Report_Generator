#### analysis/llm_interaction.py
####
#### This script is responsible for interacting with a local Language Model (LLM).
#### It provides a function to query the LLM with a given prompt and retrieve its
#### response. The script constructs a specific prompt that instructs the LLM to
#### act as a data analyst assistant, providing factual analysis based solely on
#### the provided data and avoiding speculative statements.
####
#### You may experiment tweaking 'system_prompt' below to customize the end result.


import requests
import logging


def query_llm(prompt: str, model: str, temperature: float) -> str | None:
    """Queries local LLM with data-focused prompting."""
    system_prompt = """You are a data analyst assistant. Your task is to analyze and describe 
    the provided data in a factual manner. Follow these rules:
    1. Only use information explicitly provided in the data
    2. Do not make assumptions beyond what's in the numbers
    3. Avoid speculative language like "might", "could", "possibly"
    4. State exact percentages and values from the data
    5. If no notable patterns exist, say so directly
    6. Use clear, concise business language"""

    full_prompt = f"{system_prompt}\n\n{prompt}"

    try:
        response = requests.post(
            "http://localhost:11434/api/generate",
            json={
                "model": model,
                "prompt": full_prompt,
                "temperature": temperature,
                "stream": False
            },
            timeout=120
        )
        response.raise_for_status()
        return response.json().get("response", "").strip()
    except requests.exceptions.RequestException as e:
        logging.error(f"Error querying LLM (Connection Error): {e}")
        print(f"Error querying LLM (Connection Error): {e}")
        return None
    except Exception as e:
        logging.error(f"Error querying LLM: {e}")
        print(f"Error querying LLM: {e}")
        return None