"""
updater.py - Backward-compatible facade for Auto and Manual Updates.
Re-exports updater service functions from services.updater_service.
"""

from services.updater_service import (
    DEFAULT_REPO,
    FALLBACK_VERSION,
    get_bundle_dir,
    get_current_version,
    get_github_token,
    parse_version,
    is_newer_version,
    check_for_updates,
    download_and_install_update,
)

__all__ = [
    "DEFAULT_REPO",
    "FALLBACK_VERSION",
    "get_bundle_dir",
    "get_current_version",
    "get_github_token",
    "parse_version",
    "is_newer_version",
    "check_for_updates",
    "download_and_install_update",
]
