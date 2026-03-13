"""Tests for navigation cache."""

import pytest
import json
from pathlib import Path
from datetime import datetime
from csvtool.cache import NavigationCache, CachedNavigation


class TestCachedNavigation:
    """Tests for CachedNavigation model."""

    def test_cached_navigation_creation(self):
        """Test creating a cached navigation entry."""
        cached = CachedNavigation(
            selector="button.submit",
            action="click",
            success_count=5,
            last_success=datetime.now()
        )
        assert cached.selector == "button.submit"
        assert cached.success_count == 5

    def test_increment_success(self):
        """Test incrementing success count."""
        cached = CachedNavigation(
            selector="button.submit",
            action="click",
            success_count=5,
            last_success=datetime.now()
        )
        cached.record_success()
        assert cached.success_count == 6


class TestNavigationCache:
    """Tests for NavigationCache."""

    @pytest.fixture
    def cache_dir(self, tmp_path: Path) -> Path:
        """Create temporary cache directory."""
        cache_path = tmp_path / "cache"
        cache_path.mkdir()
        return cache_path

    @pytest.fixture
    def cache(self, cache_dir: Path) -> NavigationCache:
        """Create cache instance."""
        return NavigationCache(
            cache_dir=cache_dir,
            product_url="https://app.example.com"
        )

    def test_cache_initialization(self, cache: NavigationCache):
        """Test cache initializes empty."""
        assert len(cache.navigations) == 0

    def test_cache_get_miss(self, cache: NavigationCache):
        """Test cache miss returns None."""
        result = cache.get("nonexistent instruction")
        assert result is None

    def test_cache_set_and_get(self, cache: NavigationCache):
        """Test setting and getting cache entry."""
        cache.set(
            instruction="click login button",
            selector="button.login",
            action="click"
        )

        result = cache.get("click login button")
        assert result is not None
        assert result.selector == "button.login"

    def test_cache_case_insensitive(self, cache: NavigationCache):
        """Test cache lookup is case insensitive."""
        cache.set(
            instruction="Click Login Button",
            selector="button.login",
            action="click"
        )

        result = cache.get("click login button")
        assert result is not None

    def test_cache_normalizes_whitespace(self, cache: NavigationCache):
        """Test cache normalizes whitespace in instructions."""
        cache.set(
            instruction="  click   login   button  ",
            selector="button.login",
            action="click"
        )

        result = cache.get("click login button")
        assert result is not None

    def test_cache_persists_to_file(self, cache: NavigationCache, cache_dir: Path):
        """Test cache saves to JSON file."""
        cache.set(
            instruction="click login",
            selector="button.login",
            action="click"
        )
        cache.save()

        cache_file = cache_dir / "app.example.com.json"
        assert cache_file.exists()

        data = json.loads(cache_file.read_text())
        assert "click login" in data["navigations"]

    def test_cache_loads_from_file(self, cache_dir: Path):
        """Test cache loads existing file."""
        # Create cache file
        cache_file = cache_dir / "app.example.com.json"
        cache_data = {
            "cache_version": "1.0",
            "product_url": "https://app.example.com",
            "last_updated": datetime.now().isoformat(),
            "navigations": {
                "click submit": {
                    "selector": "button[type='submit']",
                    "action": "click",
                    "success_count": 10,
                    "last_success": datetime.now().isoformat()
                }
            }
        }
        cache_file.write_text(json.dumps(cache_data))

        # Load cache
        cache = NavigationCache(
            cache_dir=cache_dir,
            product_url="https://app.example.com"
        )

        result = cache.get("click submit")
        assert result is not None
        assert result.success_count == 10

    def test_cache_update_increments_count(self, cache: NavigationCache):
        """Test updating existing cache entry."""
        cache.set("click button", "button.btn", "click")
        cache.set("click button", "button.btn", "click")

        result = cache.get("click button")
        assert result.success_count == 2

    def test_cache_update_with_new_selector(self, cache: NavigationCache):
        """Test updating cache with different selector."""
        cache.set("click button", "button.old", "click")
        cache.set("click button", "button.new", "click")

        result = cache.get("click button")
        assert result.selector == "button.new"
        assert result.success_count == 1  # Reset on selector change
