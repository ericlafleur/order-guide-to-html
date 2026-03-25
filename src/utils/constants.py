"""Constants used throughout the order guide processing."""

import re
from collections import OrderedDict
from typing import Dict, Tuple

# Status labels for equipment availability
STATUS_LABELS = OrderedDict(
    [
        ("S", "Standard Equipment"),
        ("A", "Available"),
        ("D", "ADI Available"),
        ("■", "Included in Equipment Group"),
        ("□", "Included in Equipment Group but upgradeable"),
        ("*", "Indicates availability of feature on multiple models"),
        ("--", "Not Available"),
    ]
)

# Propulsion keywords for inference
PROPULSION_KEYWORDS: Dict[str, str] = {
    'ev': 'EV',
    'electric': 'EV',
    'bev': 'EV',
    'hybrid': 'HYBRID',
    'phev': 'PHEV',
    'turbomax': 'ICE',
    'ecotec': 'ICE',
    'duramax': 'ICE',
    'turbo-diesel': 'ICE',
    'diesel': 'ICE',
    'v8': 'ICE',
    'v6': 'ICE',
    'turbo': 'ICE',
}

# Vehicle type keywords
VEHICLE_TYPE_SUV_KEYWORDS = (
    'equinox', 'blazer', 'trax', 'trailblazer', 'traverse', 'tahoe', 'suburban',
    'terrain', 'acadia', 'encore', 'envision', 'enclave', 'escalade',
    'xt4', 'xt5', 'xt6', 'lyriq', 'yukon',
)

VEHICLE_TYPE_TRUCK_KEYWORDS = (
    'silverado', 'sierra', 'colorado', 'canyon', 'titan', 'f-', 'ram ', 'tundra'
)

# Drive type patterns for regex matching
DRIVE_TYPE_PATTERNS = [
    (re.compile(r'\beawd\b', re.I), 'eAWD'),
    (re.compile(r'\bawd\b', re.I), 'AWD'),
    (re.compile(r'\bfwd\b', re.I), 'FWD'),
    (re.compile(r'\brwd\b', re.I), 'RWD'),
    (re.compile(r'\b4wd\b|\b4x4\b', re.I), '4WD'),
    (re.compile(r'\b2wd\b|\b4x2\b', re.I), '2WD'),
    (re.compile(r'\ball[- ]wheel drive\b', re.I), 'AWD'),
    (re.compile(r'\bfront[- ]wheel drive\b', re.I), 'FWD'),
    (re.compile(r'\brear[- ]wheel drive\b', re.I), 'RWD'),
    (re.compile(r'\b4[- ]wheel drive\b', re.I), '4WD'),
    (re.compile(r'\b2[- ]wheel drive\b', re.I), '2WD'),
]

# Regular expressions
FOOTNOTE_LINE_RE = re.compile(r"^\s*(\d+)\.\s*(.+?)\s*$")
URL_RE = re.compile(r"(https?://[^\s<]+)")
TRAILING_DIGITS_RE = re.compile(r"^(.*?)(\d+)\s*$")
CODE_IN_PARENS_RE = re.compile(r"\(([A-Z0-9]{2,6})\)")
NON_ALNUM_RE = re.compile(r"[^a-z0-9]+", re.I)

# Filler tokens to remove from model names
FILLER_TOKENS = {"Truck", "Trucks", "Cars", "Car", "SUV", "SUVs"}

# Status priority for sorting
STATUS_PRIORITY = {
    'Standard Equipment': 0,
    'Included in Equipment Group': 1,
    'Included in Equipment Group but upgradeable': 2,
    'ADI Available': 3,
    'Available': 4,
    'Indicates availability of feature on multiple models': 5,
    'Not Available': 6,
}

# Display order for trim availability sections
TRIM_SECTION_ORDER = [
    'Standard Equipment',
    'Included in Equipment Group',
    'Included in Equipment Group but upgradeable',
    'ADI Available',
    'Available',
    'Indicates availability of feature on multiple models',
]

# Category sequence for feature organization
CATEGORY_SEQUENCE = [
    'Safety and driver assistance',
    'Technology and connectivity',
    'Interior and comfort',
    'Exterior and utility',
    'Wheels and tires',
    'Mechanical and performance',
    'Packages and options',
    'Colour and trim',
    'Specifications and dimensions',
    'Engine, axle and GVWR',
    'Trailering and GCWR',
    'Other guide content',
]

CATEGORY_ORDER = {name: idx for idx, name in enumerate(CATEGORY_SEQUENCE)}

# Manifest-specific constants
MANIFEST_STANDARDISH_CODES = {'S', '■', '□'}
MANIFEST_DRIVE_TOKENS = ('2WD', '4WD', 'AWD', 'FWD', 'RWD')
MANIFEST_BODY_STYLE_TOKENS = ('Crew Cab', 'Double Cab', 'Regular Cab')
