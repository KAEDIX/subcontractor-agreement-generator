# backend/khp_registry.py
"""
Loads KHP property identity/address data from the khp-property-cache repo
(source of truth for KHP project names and addresses — see that repo's
README). Directory: KHP_CACHE_DIR env var, default ~/GitHub/khp-property-cache.
Each property lives at KHP00X/khp00x.json with an "identity" block
(khp_code, short_name, street_address, city, state, zip).
"""

import glob
import json
import os


class KHPRegistryError(Exception):
    """Raised when the khp-property-cache directory is missing or unreadable."""


def _cache_dir():
    return os.path.expanduser(os.getenv("KHP_CACHE_DIR", "~/GitHub/khp-property-cache"))


# Lifecycle states that mean "not a live KAEDIX project". Compared against the
# WHOLE field value, never as a substring: KHP003's disposition prose contains
# the phrase "a dead contract" while KHP003 itself is very much active.
INACTIVE_STATUSES = {"sold", "dead", "canceled", "cancelled", "killed"}


def _is_active(record):
    """False when the cache record carries an explicit dead/sold/canceled state.
    A record with no status field counts as active -- that is how KHP006 and
    KHP008 are carried, and a missing field must never silently hide a live
    project. Note `status` is sometimes a dict (a stage_line block), which is
    not a lifecycle marker."""
    status = record.get("status")
    if isinstance(status, str) and status.strip().lower() in INACTIVE_STATUSES:
        return False
    economics = record.get("deal_economics")
    if isinstance(economics, dict):
        deal_status = economics.get("deal_status")
        if isinstance(deal_status, str) and deal_status.strip().lower() in INACTIVE_STATUSES:
            return False
    return True


def load_properties(active_only=False):
    """Returns a list of {khp_code, short_name, street_address, city, state, zip}
    dicts, one per KHP property found in the cache, sorted by khp_code.
    active_only drops properties the cache marks sold/dead/canceled.
    Raises KHPRegistryError if the cache directory doesn't exist."""
    cache_dir = _cache_dir()
    if not os.path.isdir(cache_dir):
        raise KHPRegistryError(f"khp-property-cache not found at {cache_dir}")

    properties = []
    for path in sorted(glob.glob(os.path.join(cache_dir, "KHP*", "khp*.json"))):
        try:
            with open(path) as f:
                data = json.load(f)
        except (OSError, json.JSONDecodeError):
            continue
        identity = data.get("identity", data)
        khp_code = identity.get("khp_code")
        if not khp_code:
            continue
        if active_only and not _is_active(data):
            continue
        properties.append({
            "khp_code": khp_code,
            "short_name": identity.get("short_name", ""),
            "street_address": identity.get("street_address", ""),
            "city": identity.get("city", ""),
            "state": identity.get("state", ""),
            "zip": identity.get("zip", ""),
        })
    properties.sort(key=lambda p: p["khp_code"])
    return properties
