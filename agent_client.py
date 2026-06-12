from __future__ import annotations

import os
from pathlib import Path
from typing import List, Dict, Optional

from dotenv import load_dotenv
from azure.identity import DefaultAzureCredential, get_bearer_token_provider
from openai import OpenAI


class InvoiceIQAgentClient:
    """Small Microsoft Foundry Responses API client for Invoice IQ Agent."""

    def __init__(self, endpoint: Optional[str] = None):
        env_path = Path(__file__).with_name(".env")
        load_dotenv(dotenv_path=env_path)

        self.agent_endpoint = endpoint or os.getenv("AGENT_ENDPOINT")
        if not self.agent_endpoint:
            raise ValueError("Missing AGENT_ENDPOINT in .env file.")

        self.agent_endpoint = self.agent_endpoint.strip().rstrip("/")
        if self.agent_endpoint.endswith("/responses"):
            self.agent_endpoint = self.agent_endpoint.removesuffix("/responses")

        self.client = OpenAI(
            api_key=get_bearer_token_provider(
                DefaultAzureCredential(),
                "https://ai.azure.com/.default",
            ),
            base_url=self.agent_endpoint,
            default_query={"api-version": "v1"},
        )

        self.conversation_history: List[Dict[str, str]] = []

    def send_message(self, user_message: str, use_history: bool = True) -> str:
        if use_history:
            self.conversation_history.append({"role": "user", "content": user_message})
            response = self.client.responses.create(input=self.conversation_history)
            assistant_message = response.output_text
            self.conversation_history.append({"role": "assistant", "content": assistant_message})
            return assistant_message

        response = self.client.responses.create(
            input=[{"role": "user", "content": user_message}]
        )
        return response.output_text


if __name__ == "__main__":
    agent = InvoiceIQAgentClient()
    print("Invoice IQ Agent is ready. Type 'quit' to exit.")
    while True:
        prompt = input("\nYou: ").strip()
        if prompt.lower() == "quit":
            break
        print("\nAgent:", agent.send_message(prompt))
