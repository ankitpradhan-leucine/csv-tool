"""Hybrid AI navigator supporting both local LLM (Ollama) and Claude API."""

import json
import logging
import os
import requests
from typing import Optional

from csvtool.ai_navigator import AINavigator, NavigationAction, ActionType

logger = logging.getLogger(__name__)


class HybridNavigator:
    """Hybrid navigator that tries local LLM first, falls back to Claude."""
    
    def __init__(self, api_key: Optional[str] = None, ollama_url: str = "http://localhost:11434", ollama_model: str = "llava:13b"):
        """Initialize hybrid navigator."""
        self.ollama_url = ollama_url
        self.ollama_model = ollama_model
        self.claude_navigator = AINavigator(api_key=api_key)
        
        # Statistics tracking
        self.local_success_count = 0
        self.local_failure_count = 0
        self.claude_count = 0
        
        # Check if Ollama is available
        self.ollama_available = self._check_ollama()
        
        if self.ollama_available:
            logger.info(f"🚀 Hybrid Mode Enabled | Local LLM: {ollama_model} @ {ollama_url} | Fallback: Claude Sonnet 4")
        else:
            logger.warning(f"⚠️ Local LLM unavailable at {ollama_url} - using Claude only")
    
    def _check_ollama(self) -> bool:
        """Check if Ollama server is running."""
        try:
            response = requests.get(f"{self.ollama_url}/api/tags", timeout=2)
            return response.status_code == 200
        except Exception:
            return False

    @property
    def total_input_tokens(self) -> int:
        """Get total input tokens from Claude API (local LLM doesn't use tokens)."""
        return self.claude_navigator.total_input_tokens

    @property
    def total_output_tokens(self) -> int:
        """Get total output tokens from Claude API (local LLM doesn't use tokens)."""
        return self.claude_navigator.total_output_tokens
    
    async def analyze(self, screenshot_base64: str, instruction: str, context: Optional[str] = None) -> NavigationAction:
        """Analyze screenshot - try local first, fallback to Claude."""
        
        if self.ollama_available:
            try:
                logger.info(f"🤖 Trying LOCAL LLM ({self.ollama_model}) for: {instruction[:50]}...")
                result = self._analyze_local(screenshot_base64, instruction, context)
                
                # Check confidence
                confidence = result.get("confidence", 0.5)
                
                if confidence >= 0.7:  # High confidence threshold
                    logger.info(f"✅ LOCAL LLM SUCCESS | Confidence: {confidence:.0%} | Selector: {result.get('selector', 'N/A')}")
                    self.local_success_count += 1
                    
                    return NavigationAction(
                        action_type=ActionType(result.get("action_type", "click")),
                        selector=result.get("selector"),
                        text=result.get("text"),
                        description=result.get("description", "")
                    )
                else:
                    logger.warning(f"⚠️ LOCAL LLM LOW CONFIDENCE ({confidence:.0%}) - falling back to Claude")
                    self.local_failure_count += 1
                    
            except Exception as e:
                logger.warning(f"⚠️ LOCAL LLM FAILED: {str(e)} - falling back to Claude")
                self.local_failure_count += 1
        
        # Fallback to Claude
        logger.info(f"☁️ Using CLAUDE for: {instruction[:50]}...")
        self.claude_count += 1
        return await self.claude_navigator.analyze(screenshot_base64, instruction, context)
    
    def _analyze_local(self, screenshot_base64: str, instruction: str, context: Optional[str] = None) -> dict:
        """Analyze using local Ollama LLM."""
        
        prompt = f"""You are a browser automation assistant. Analyze the screenshot and determine the action.

Instruction: {instruction}
Context: {context or 'None'}

Return ONLY valid JSON with this EXACT format:
{{
  "action_type": "click",
  "selector": "button >> text=Login",
  "text": null,
  "description": "Click the login button",
  "confidence": 0.9
}}

Rules:
- action_type: "click", "type", "wait", "navigate", or "verify"
- selector: Use Playwright selectors (CSS, text=, [attr=])
- confidence: 0.0 to 1.0 (how sure you are this is correct)
- Return ONLY the JSON, no explanation

Valid selector examples:
- "button >> text=Log In"
- "input[type='email']"
- "text=Submit"
- "#login-button"
"""

        response = requests.post(
            f"{self.ollama_url}/api/generate",
            json={
                "model": self.ollama_model,
                "prompt": prompt,
                "images": [screenshot_base64],
                "stream": False
            },
            timeout=30
        )
        
        result = response.json()
        text = result.get('response', '')
        
        # Extract JSON
        if '```json' in text:
            start = text.find('```json') + 7
            end = text.find('```', start)
            text = text[start:end].strip()
        elif '```' in text:
            text = text.split('```')[1]
            if text.startswith('json'):
                text = text[4:]
        elif '{' in text:
            text = text[text.find('{'):text.rfind('}')+1]
        
        return json.loads(text)
    
    async def verify_result(self, screenshot_base64: str, expected_result: str) -> dict:
        """Verify result - try local first, fallback to Claude."""
        
        if self.ollama_available:
            try:
                logger.info(f"🤖 Trying LOCAL LLM for verification...")
                result = self._verify_local(screenshot_base64, expected_result)
                
                confidence = result.get("confidence", 0.5)
                
                if confidence >= 0.7:
                    logger.info(f"✅ LOCAL LLM VERIFICATION | Confidence: {confidence:.0%} | Passed: {result.get('passed')}")
                    self.local_success_count += 1
                    return result
                else:
                    logger.warning(f"⚠️ LOCAL LLM LOW CONFIDENCE ({confidence:.0%}) - falling back to Claude")
                    self.local_failure_count += 1
                    
            except Exception as e:
                logger.warning(f"⚠️ LOCAL LLM FAILED: {str(e)} - falling back to Claude")
                self.local_failure_count += 1
        
        # Fallback to Claude
        logger.info(f"☁️ Using CLAUDE for verification")
        self.claude_count += 1
        return await self.claude_navigator.verify_result(screenshot_base64, expected_result)
    
    def _verify_local(self, screenshot_base64: str, expected_result: str) -> dict:
        """Verify using local Ollama LLM."""
        
        prompt = f"""Compare the screenshot against the expected result.

Expected Result: {expected_result}

Return ONLY valid JSON:
{{
  "passed": true,
  "observed_result": "what you see on screen",
  "reasoning": "brief explanation",
  "confidence": 0.9
}}

confidence: 0.0 to 1.0 (how sure you are about this verification)
"""

        response = requests.post(
            f"{self.ollama_url}/api/generate",
            json={
                "model": self.ollama_model,
                "prompt": prompt,
                "images": [screenshot_base64],
                "stream": False
            },
            timeout=30
        )
        
        result = response.json()
        text = result.get('response', '')
        
        # Extract JSON
        if '```json' in text:
            start = text.find('```json') + 7
            end = text.find('```', start)
            text = text[start:end].strip()
        elif '{' in text:
            text = text[text.find('{'):text.rfind('}')+1]
        
        return json.loads(text)
    
    async def identify_login_form(self, screenshot_base64: str) -> dict:
        """Identify login form - always use Claude for reliability."""
        logger.info(f"☁️ Using CLAUDE for login form identification (requires high accuracy)")
        self.claude_count += 1
        return await self.claude_navigator.identify_login_form(screenshot_base64)
    
    def log_statistics(self):
        """Log usage statistics."""
        total = self.local_success_count + self.local_failure_count + self.claude_count

        if total == 0:
            return

        local_success_pct = (self.local_success_count / total * 100) if total > 0 else 0
        claude_pct = (self.claude_count / total * 100) if total > 0 else 0

        logger.info("🤖 Hybrid LLM Breakdown:")
        logger.info(f"  Local LLM Success: {self.local_success_count} ({local_success_pct:.1f}%)")
        logger.info(f"  Local LLM Fallback: {self.local_failure_count}")
        logger.info(f"  Claude API Calls: {self.claude_count} ({claude_pct:.1f}%)")
        logger.info(f"  Total LLM Calls: {total}")
