import os
from langchain.prompts import PromptTemplate

USE_LOCAL_AI = os.getenv("USE_LOCAL_AI", "True") == "True"

if USE_LOCAL_AI:
    from langchain_community.llms import Ollama
else:
    from langchain.chat_models import ChatOpenAI


def summarize_excel_with_ollama(excel_data):
    # Set up the model
    if USE_LOCAL_AI:
        llm = Ollama(model="llama3", base_url="http://127.0.0.1:11434")
    else:
        llm = ChatOpenAI(temperature=0.3, model_name="gpt-3.5-turbo")

    # Prompt template
    template = """
    You are a helpful assistant. Summarize the following Excel data and provide key observations or suggestions.

    Excel Data:
    {excel_data}

    Summary:
    """
    prompt = PromptTemplate(input_variables=["excel_data"], template=template)

    # Preview data (to avoid hitting token limits, adjust accordingly)
    preview_data = excel_data.to_string(index=False)  # You can change or remove this limit

    # Format prompt with preview data
    formatted_prompt = prompt.format(excel_data=preview_data)

    # Get model response
    response = llm.invoke(formatted_prompt)

    # Ensure to return the content from the response
    if isinstance(response, dict):  # Check for dictionary-based response structure
        return response.get("content", "")
    return response  # If the response is direct content

