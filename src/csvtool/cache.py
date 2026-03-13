"""Navigation cache for storing learned selectors."""

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Optional
from urllib.parse import urlparse

from pydantic import BaseModel, Field


class CachedNavigation(BaseModel):
    """A cached navigation entry."""

    selector: str
    action: str
    success_count: int = 1
    last_success: datetime = Field(default_factory=datetime.now)

    def record_success(self) -> None:
        """Record a successful use of this cached navigation."""
        self.success_count += 1
        self.last_success = datetime.now()


class NavigationCache:
    """Cache for storing learned navigation selectors."""

    CACHE_VERSION = "1.0"

    def __init__(self, cache_dir: Path, product_url: str):
        """Initialize cache for a specific product URL."""
        self.cache_dir = Path(cache_dir)
        self.product_url = product_url
        self.navigations: dict[str, CachedNavigation] = {}

        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self._load()

    @property
    def cache_file(self) -> Path:
        """Get cache file path based on product URL."""
        # Extract domain from URL for filename
        parsed = urlparse(self.product_url)
        domain = parsed.netloc or parsed.path
        # Sanitize for filename
        safe_name = re.sub(r"[^\w\-.]", "_", domain)
        return self.cache_dir / f"{safe_name}.json"

    @staticmethod
    def _normalize_instruction(instruction: str) -> str:
        """Normalize instruction for cache key."""
        # Lowercase, collapse whitespace, strip
        normalized = instruction.lower().strip()
        normalized = re.sub(r"\s+", " ", normalized)
        return normalized

    def get(self, instruction: str) -> Optional[CachedNavigation]:
        """Get cached navigation for instruction."""
        key = self._normalize_instruction(instruction)
        return self.navigations.get(key)

    def set(
        self,
        instruction: str,
        selector: str,
        action: str
    ) -> None:
        """Set or update cached navigation."""
        key = self._normalize_instruction(instruction)

        existing = self.navigations.get(key)
        if existing and existing.selector == selector:
            # Same selector - increment success count
            existing.record_success()
        else:
            # New entry or different selector
            self.navigations[key] = CachedNavigation(
                selector=selector,
                action=action,
                success_count=1,
                last_success=datetime.now()
            )

    def save(self) -> None:
        """Save cache to file."""
        data = {
            "cache_version": self.CACHE_VERSION,
            "product_url": self.product_url,
            "last_updated": datetime.now().isoformat(),
            "navigations": {
                key: {
                    "selector": nav.selector,
                    "action": nav.action,
                    "success_count": nav.success_count,
                    "last_success": nav.last_success.isoformat()
                }
                for key, nav in self.navigations.items()
            }
        }
        self.cache_file.write_text(json.dumps(data, indent=2))

    def _load(self) -> None:
        """Load cache from file if exists."""
        if not self.cache_file.exists():
            return

        try:
            data = json.loads(self.cache_file.read_text())

            for key, nav_data in data.get("navigations", {}).items():
                self.navigations[key] = CachedNavigation(
                    selector=nav_data["selector"],
                    action=nav_data["action"],
                    success_count=nav_data.get("success_count", 1),
                    last_success=datetime.fromisoformat(
                        nav_data.get("last_success", datetime.now().isoformat())
                    )
                )
        except (json.JSONDecodeError, KeyError) as e:
            # Invalid cache file - start fresh
            self.navigations = {}
