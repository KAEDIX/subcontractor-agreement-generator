# backend/khp_registry.py
"""
Loads KHP property identity/address data from the khp-property-cache repo
(source of truth for KHP project names and addresses — see that repo's
README). Directory: KHP_CACHE_DIR env var, default ~/GitHub/khp-property-cache.
Each property lives at KHP00X/khp00x.json with an "identity" block
(khp_code, short_name, street_address, city, state, zip) and a top-level
"status" field (e.g. "active-construction", "sold", "dead").
"""

import glob
import json
import os


class KHPRegistryError(Exception):
    """Raised when the khp-property-cache directory is missing or unreadable."""


def _cache_dir():
    return os.path.expanduser(os.getenv("KHP_CACHE_DIR", "~/GitHub/khp-property-cache"))


def load_properties():
    """Returns a list of {khp_code, short_name, street_address, city, state,
    zip, status} dicts, one per KHP property found in the cache, sorted by
    khp_code. status is the record's top-level "status" field, or None when
    the record doesn't set one. Raises KHPRegistryError if the cache
    directory doesn't exist."""
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
        properties.append({
            "khp_code": khp_code,
            "short_name": identity.get("short_name", ""),
            "street_address": identity.get("street_address", ""),
            "city": identity.get("city", ""),
            "state": identity.get("state", ""),
            "zip": identity.get("zip", ""),
            "status": data.get("status"),
        })
    properties.sort(key=lambda p: p["khp_code"])
    return properties
