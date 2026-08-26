# -*- coding: utf-8 -*-
"""
Handles loading and saving the per-item prices to a local JSON file
(prices.json) so that prices set by the shop owner persist between runs.
"""

import json
import os

PRICES_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "prices.json")


def load_prices(default_keys):
    """Load prices.json. Any missing item key gets a default price of 0."""
    prices = {key: 0 for key in default_keys}
    if os.path.exists(PRICES_FILE):
        try:
            with open(PRICES_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            for key in default_keys:
                if key in saved:
                    prices[key] = saved[key]
        except (json.JSONDecodeError, OSError):
            pass
    return prices


def save_prices(prices):
    with open(PRICES_FILE, "w", encoding="utf-8") as f:
        json.dump(prices, f, ensure_ascii=False, indent=2)
