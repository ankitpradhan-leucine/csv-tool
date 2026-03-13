"""Tests for AI navigator using Claude."""

import pytest
from unittest.mock import Mock, AsyncMock, patch, MagicMock
from csvtool.ai_navigator import AINavigator, NavigationAction, ActionType


class TestNavigationAction:
    """Tests for NavigationAction model."""

    def test_click_action(self):
        """Test creating a click action."""
        action = NavigationAction(
            action_type=ActionType.CLICK,
            selector="button.submit",
            description="Click submit button"
        )
        assert action.action_type == ActionType.CLICK
        assert action.selector == "button.submit"

    def test_type_action(self):
        """Test creating a type action."""
        action = NavigationAction(
            action_type=ActionType.TYPE,
            selector="input[name='email']",
            text="test@example.com",
            description="Enter email"
        )
        assert action.action_type == ActionType.TYPE
        assert action.text == "test@example.com"

    def test_verify_action(self):
        """Test creating a verify action."""
        action = NavigationAction(
            action_type=ActionType.VERIFY,
            expected_text="Success",
            description="Verify success message"
        )
        assert action.action_type == ActionType.VERIFY


class TestAINavigator:
    """Tests for AINavigator class."""

    @pytest.fixture
    def mock_anthropic_client(self):
        """Create mock Anthropic client."""
        with patch("csvtool.ai_navigator.anthropic.Anthropic") as mock:
            client = MagicMock()
            mock.return_value = client
            yield client

    @pytest.fixture
    def navigator(self, mock_anthropic_client):
        """Create AINavigator with mocked client."""
        return AINavigator(api_key="test-key")

    def test_navigator_initialization(self, mock_anthropic_client):
        """Test navigator initializes with API key."""
        navigator = AINavigator(api_key="test-key")
        assert navigator is not None

    def test_navigator_requires_api_key(self):
        """Test navigator raises error without API key."""
        with patch.dict("os.environ", {}, clear=True):
            with pytest.raises(ValueError, match="ANTHROPIC_API_KEY"):
                AINavigator()

    @pytest.mark.asyncio
    async def test_analyze_returns_action(self, navigator, mock_anthropic_client):
        """Test analyze returns a NavigationAction."""
        # Mock Claude response
        mock_response = MagicMock()
        mock_response.content = [MagicMock(text='''
        {
            "action_type": "click",
            "selector": "button.login",
            "description": "Click login button"
        }
        ''')]
        mock_anthropic_client.messages.create.return_value = mock_response

        action = await navigator.analyze(
            screenshot_base64="base64data",
            instruction="Click the login button"
        )

        assert isinstance(action, NavigationAction)
        assert action.action_type == ActionType.CLICK

    @pytest.mark.asyncio
    async def test_verify_result(self, navigator, mock_anthropic_client):
        """Test verifying expected result against screenshot."""
        mock_response = MagicMock()
        mock_response.content = [MagicMock(text='''
        {
            "passed": true,
            "observed_result": "Success message displayed",
            "reasoning": "The page shows 'Operation successful'"
        }
        ''')]
        mock_anthropic_client.messages.create.return_value = mock_response

        result = await navigator.verify_result(
            screenshot_base64="base64data",
            expected_result="Success message displayed"
        )

        assert result["passed"] is True
        assert "observed_result" in result

    @pytest.mark.asyncio
    async def test_generate_test_data(self, navigator, mock_anthropic_client):
        """Test generating test data for form fields."""
        mock_response = MagicMock()
        mock_response.content = [MagicMock(text='''
        {
            "fields": [
                {"selector": "input[name='name']", "value": "John Doe"},
                {"selector": "input[name='email']", "value": "john@example.com"}
            ]
        }
        ''')]
        mock_anthropic_client.messages.create.return_value = mock_response

        data = await navigator.generate_test_data(
            screenshot_base64="base64data",
            instruction="Enter details",
            skip_mandatory=False
        )

        assert len(data["fields"]) == 2

    @pytest.mark.asyncio
    async def test_identify_login_form(self, navigator, mock_anthropic_client):
        """Test identifying login form elements."""
        mock_response = MagicMock()
        mock_response.content = [MagicMock(text='''
        {
            "username_selector": "input[name='username']",
            "password_selector": "input[type='password']",
            "submit_selector": "button[type='submit']",
            "is_login_page": true
        }
        ''')]
        mock_anthropic_client.messages.create.return_value = mock_response

        result = await navigator.identify_login_form(screenshot_base64="base64data")

        assert result["is_login_page"] is True
        assert "username_selector" in result
        assert "password_selector" in result
