"""AI-powered navigation using Claude API."""

import json
import os
from enum import Enum
from typing import Optional

import anthropic
from pydantic import BaseModel, Field


class ActionType(str, Enum):
    """Types of browser actions."""
    CLICK = "click"
    TYPE = "type"
    VERIFY = "verify"
    WAIT = "wait"
    NAVIGATE = "navigate"


class NavigationAction(BaseModel):
    """An action to perform in the browser."""

    action_type: ActionType
    selector: Optional[str] = None
    text: Optional[str] = None
    expected_text: Optional[str] = None
    url: Optional[str] = None
    description: str = ""


class AINavigator:
    """AI-powered navigation using Claude Vision API."""

    MODEL = "claude-sonnet-4-20250514"
    MAX_TOKENS = 1024

    def __init__(self, api_key: Optional[str] = None):
        """Initialize with Anthropic API key."""
        self.api_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
        if not self.api_key:
            raise ValueError("ANTHROPIC_API_KEY must be provided or set in environment")
        self.client = anthropic.Anthropic(api_key=self.api_key)

    async def analyze(
        self,
        screenshot_base64: str,
        instruction: str,
        context: Optional[str] = None
    ) -> NavigationAction:
        """Analyze screenshot and determine next action for instruction."""
        system_prompt = """You are a browser automation assistant. Analyze the screenshot and determine the exact action needed to fulfill the instruction.

IMPORTANT: For selectors, use ONLY these valid Playwright selector formats:
1. Standard CSS selectors: "button.login", "#username", "input[type='password']"
2. Text selectors: "text=Log In", "text=Submit" (for elements containing exact text)
3. Attribute selectors: "[data-testid='login-button']", "[name='email']"
4. Combination: "button >> text=Log In" (button containing text)

DO NOT use jQuery selectors like ":contains()" - they are INVALID.
PREFER text selectors when identifying buttons by their label.

Return a JSON object with:
- action_type: "click", "type", "verify", "wait", or "navigate"
- selector: Valid Playwright selector (for click/type actions)
- text: text to type (for type actions)
- expected_text: text to verify (for verify actions)
- description: brief description of the action

Return ONLY valid JSON, no markdown or explanation."""

        user_content = [
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": screenshot_base64
                }
            },
            {
                "type": "text",
                "text": f"Instruction: {instruction}\n\nContext: {context or 'None'}\n\nWhat action should be taken?"
            }
        ]

        response = self.client.messages.create(
            model=self.MODEL,
            max_tokens=self.MAX_TOKENS,
            system=system_prompt,
            messages=[{"role": "user", "content": user_content}]
        )

        response_text = response.content[0].text.strip()

        # Parse JSON response
        try:
            # Handle potential markdown code blocks and explanatory text
            if "```json" in response_text:
                # Extract JSON from code block
                start = response_text.find("```json") + 7
                end = response_text.find("```", start)
                response_text = response_text[start:end].strip()
            elif "```" in response_text:
                # Generic code block
                response_text = response_text.split("```")[1]
                if response_text.startswith("json"):
                    response_text = response_text[4:]
            elif "{" in response_text:
                # Find the first { and assume JSON starts there
                response_text = response_text[response_text.find("{"):]
                # Find the last } and cut there
                response_text = response_text[:response_text.rfind("}") + 1]

            data = json.loads(response_text)
            return NavigationAction(
                action_type=ActionType(data.get("action_type", "click")),
                selector=data.get("selector"),
                text=data.get("text"),
                expected_text=data.get("expected_text"),
                description=data.get("description", "")
            )
        except (json.JSONDecodeError, KeyError) as e:
            raise ValueError(f"Failed to parse AI response: {response_text}") from e

    async def verify_result(
        self,
        screenshot_base64: str,
        expected_result: str
    ) -> dict:
        """Verify if the expected result matches the current screen state."""
        system_prompt = """You are a test verification assistant. Compare the screenshot against the expected result.

Return a JSON object with:
- passed: true/false
- observed_result: what you actually see on the screen
- reasoning: brief explanation of your verification

Return ONLY valid JSON, no markdown or explanation."""

        user_content = [
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": screenshot_base64
                }
            },
            {
                "type": "text",
                "text": f"Expected Result: {expected_result}\n\nDoes the screenshot match the expected result?"
            }
        ]

        response = self.client.messages.create(
            model=self.MODEL,
            max_tokens=self.MAX_TOKENS,
            system=system_prompt,
            messages=[{"role": "user", "content": user_content}]
        )

        response_text = response.content[0].text.strip()

        try:
            # Handle potential markdown code blocks and explanatory text
            if "```json" in response_text:
                start = response_text.find("```json") + 7
                end = response_text.find("```", start)
                response_text = response_text[start:end].strip()
            elif "```" in response_text:
                response_text = response_text.split("```")[1]
                if response_text.startswith("json"):
                    response_text = response_text[4:]
            elif "{" in response_text:
                response_text = response_text[response_text.find("{"):]
                response_text = response_text[:response_text.rfind("}") + 1]

            return json.loads(response_text)
        except json.JSONDecodeError as e:
            raise ValueError(f"Failed to parse verification response: {response_text}") from e

    async def generate_test_data(
        self,
        screenshot_base64: str,
        instruction: str,
        skip_mandatory: bool = False
    ) -> dict:
        """Generate appropriate test data for form fields visible on screen."""
        skip_instruction = ""
        if skip_mandatory:
            skip_instruction = "IMPORTANT: Identify one mandatory/required field and leave it empty (return empty string for its value)."

        system_prompt = f"""You are a test data generator. Analyze the form fields visible in the screenshot and generate appropriate test data.

{skip_instruction}

IMPORTANT: Use ONLY valid Playwright selectors:
- Standard CSS: "input[name='email']", "#username", ".form-control"
- Attribute selectors: "[placeholder='Email']", "[type='text']"
- NO jQuery selectors like ":contains()" - they are INVALID

Return a JSON object with:
- fields: array of objects with "selector" (valid Playwright selector) and "value" (data to enter)

Generate realistic dummy data appropriate for each field type (names, emails, dates, numbers, etc.).
Return ONLY valid JSON, no markdown or explanation."""

        user_content = [
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": screenshot_base64
                }
            },
            {
                "type": "text",
                "text": f"Instruction: {instruction}\n\nGenerate test data for the visible form fields."
            }
        ]

        response = self.client.messages.create(
            model=self.MODEL,
            max_tokens=self.MAX_TOKENS,
            system=system_prompt,
            messages=[{"role": "user", "content": user_content}]
        )

        response_text = response.content[0].text.strip()

        try:
            # Handle potential markdown code blocks and explanatory text
            if "```json" in response_text:
                start = response_text.find("```json") + 7
                end = response_text.find("```", start)
                response_text = response_text[start:end].strip()
            elif "```" in response_text:
                response_text = response_text.split("```")[1]
                if response_text.startswith("json"):
                    response_text = response_text[4:]
            elif "{" in response_text:
                response_text = response_text[response_text.find("{"):]
                response_text = response_text[:response_text.rfind("}") + 1]

            return json.loads(response_text)
        except json.JSONDecodeError as e:
            raise ValueError(f"Failed to parse test data response: {response_text}") from e

    async def identify_login_form(self, screenshot_base64: str) -> dict:
        """Identify login form fields on the page."""
        system_prompt = """You are a login form analyzer. Identify the username/email and password fields on the login page.

IMPORTANT: Use ONLY valid Playwright selectors:
- Standard CSS: "input[type='email']", "#password", "button[type='submit']"
- Text selectors for buttons: "text=Log In", "text=Sign In", "text=Submit"
- Attribute selectors: "[name='username']", "[placeholder='Password']"
- NO jQuery selectors like ":contains()" - they are INVALID

Return a JSON object with:
- username_selector: Valid Playwright selector for username/email field
- password_selector: Valid Playwright selector for password field
- submit_selector: Valid Playwright selector for login/submit button
- is_login_page: true/false

Return ONLY valid JSON, no markdown or explanation."""

        user_content = [
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": screenshot_base64
                }
            },
            {
                "type": "text",
                "text": "Identify the login form elements on this page."
            }
        ]

        response = self.client.messages.create(
            model=self.MODEL,
            max_tokens=self.MAX_TOKENS,
            system=system_prompt,
            messages=[{"role": "user", "content": user_content}]
        )

        response_text = response.content[0].text.strip()

        try:
            # Handle potential markdown code blocks and explanatory text
            if "```json" in response_text:
                start = response_text.find("```json") + 7
                end = response_text.find("```", start)
                response_text = response_text[start:end].strip()
            elif "```" in response_text:
                response_text = response_text.split("```")[1]
                if response_text.startswith("json"):
                    response_text = response_text[4:]
            elif "{" in response_text:
                response_text = response_text[response_text.find("{"):]
                response_text = response_text[:response_text.rfind("}") + 1]

            return json.loads(response_text)
        except json.JSONDecodeError as e:
            raise ValueError(f"Failed to parse login form response: {response_text}") from e
