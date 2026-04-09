#!/usr/bin/env python3
"""Toyo Restaurant Schedule Maker - Desktop Application"""

import json
import os
import sys
import copy
import random
from datetime import datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
except ImportError:
    messagebox.showerror("Missing Dependency",
                         "openpyxl is required.\nRun setup.sh first.")
    sys.exit(1)

try:
    from fpdf import FPDF
except ImportError:
    FPDF = None  # PDF export disabled but app still works

try:
    import cv2
    import numpy as np
    from difflib import SequenceMatcher
    # OCR runs in a separate Python 3.12 venv via ocr_worker.py subprocess
    _OCR_VENV_PYTHON = Path.home() / ".toyo_scheduler" / "ocr_venv" / "bin" / "python3"
    _OCR_WORKER = Path(__file__).parent / "ocr_worker.py"
    HAS_OCR = _OCR_VENV_PYTHON.exists() and _OCR_WORKER.exists()
except ImportError:
    HAS_OCR = False

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
DAYS = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
SHIFTS = ["morning", "mid", "night"]

MORNING_SIDEWORK = ["A", "B", "C", "D"]
MORNING_EXTRA_ORDER = ["C", "A", "B"]  # doubles in this order
NIGHT_SIDEWORK = ["A", "B", "C", "D", "E", "F", "G", "H"]
NIGHT_EXTRA_ORDER = ["F", "H", "G", "E", "D", "C", "B", "A"]  # doubles in this order

LUNCH_SERVER_TIME = "(10:15-2:15)"
LUNCH_HOST_TIMES = ["(10:15-4:15)", "(10:15-2:15)"]
LUNCH_MANAGER_TIME = "(10:15-4:15) "
MIDSHIFT_TIME = "2:00-4:30pm"

DINNER_SERVER_TIME = "(4:00 - close)"
DINNER_HOST_TIMES_WEEKDAY = ["(4:15-8:15)", "(4:15-9:15)", "(4:15-9:45)"]
DINNER_HOST_TIMES_FRIDAY = ["(4:15-9:15)", "(4:15-9:45)", "( 4:15-10:15)"]
DINNER_HOST_TIMES_WEEKEND = ["(4:15-9:15)", "(4:15-10:15)", "( 4:15-10:15)"]
DINNER_HOST_TIMES_SUNDAY = ["(2:15 -:8:45)", "(3:15-9:15)", "(4:15-9:45)"]
DINNER_MANAGER_TIME_WEEKDAY = "(4:30-10:00)"
DINNER_MANAGER_TIME_WEEKEND = "(4:15-10:30)"

APP_DIR = Path(__file__).parent
DATA_DIR = APP_DIR / "toyo_data"

# ---------------------------------------------------------------------------
# Default Staff Roster
# ---------------------------------------------------------------------------
DEFAULT_STAFF = [
    # Servers
    {"name": "Dian", "roles": ["server"], "seniority": 10, "fixed_schedule": True,
     "active": True, "flags": [], "role_preference": "server",
     "default_off": ["MON", "TUE"]},
    {"name": "Leony", "roles": ["server"], "seniority": 10, "fixed_schedule": True,
     "active": True, "flags": [], "role_preference": "server",
     "default_off": ["WED", "THU"]},
    {"name": "Will", "roles": ["server"], "seniority": 5, "fixed_schedule": False,
     "active": True, "flags": ["always_hibachi"], "role_preference": "server"},
    {"name": "Andy", "roles": ["server"], "seniority": 6, "fixed_schedule": False,
     "active": True, "flags": ["fill_in"], "role_preference": "server",
     "preferred_off": ["TUE", "WED", "THU"]},
    {"name": "Winnie", "roles": ["server", "host"], "seniority": 3, "fixed_schedule": False,
     "active": True, "flags": ["emergency_only"], "role_preference": None},
    {"name": "Cindy", "roles": ["server"], "seniority": 3, "fixed_schedule": False,
     "active": True, "flags": ["fill_in"], "role_preference": "server"},
    {"name": "Maddie", "roles": ["server"], "seniority": 5, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "server"},
    {"name": "Owen", "roles": ["server"], "seniority": 5, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "server"},
    {"name": "Sadie", "roles": ["server"], "seniority": 5, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "server"},
    {"name": "Arabella", "roles": ["server"], "seniority": 4, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "server"},
    {"name": "Kat", "roles": ["server"], "seniority": 5, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "server"},
    {"name": "Trish", "roles": ["server"], "seniority": 5, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "server"},
    {"name": "Sam", "roles": ["server"], "seniority": 4, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "server"},
    {"name": "Garret", "roles": ["server"], "seniority": 4, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "server"},
    {"name": "Collin", "roles": ["server"], "seniority": 5, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "server"},
    {"name": "Q", "roles": ["server"], "seniority": 4, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "server"},
    {"name": "Bryan", "roles": ["server"], "seniority": 4, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "server"},
    # Hosts
    {"name": "Maria", "roles": ["host"], "seniority": 5, "fixed_schedule": False,
     "active": True, "flags": ["no_closing"], "role_preference": "host"},
    {"name": "Makayla", "roles": ["host"], "seniority": 5, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "host"},
    {"name": "Catie Grey", "roles": ["host"], "seniority": 4, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "host"},
    {"name": "Olivia", "roles": ["host"], "seniority": 8, "fixed_schedule": False,
     "active": True, "flags": ["seniority_priority"], "role_preference": "host"},
    {"name": "Lilly", "roles": ["host"], "seniority": 4, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "host"},
    {"name": "Ashlyn", "roles": ["host"], "seniority": 4, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "host"},
    {"name": "Addison", "roles": ["host"], "seniority": 4, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "host"},
    {"name": "Hannah", "roles": ["host"], "seniority": 4, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "host"},
    # Dual role
    {"name": "Jazz", "roles": ["server", "host"], "seniority": 5, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "server"},
    {"name": "Kamryn", "roles": ["server", "host"], "seniority": 5, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "host"},
    {"name": "Abigail", "roles": ["server", "host"], "seniority": 5, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "server"},
    # Emergency / special
    {"name": "Ross", "roles": ["server", "host"], "seniority": 3, "fixed_schedule": False,
     "active": True, "flags": ["emergency_only"], "role_preference": None},
    {"name": "Shayne", "roles": ["server"], "seniority": 3, "fixed_schedule": False,
     "active": True, "flags": ["emergency_only"], "role_preference": "server"},
    # Managers
    {"name": "Aaron", "roles": ["manager"], "seniority": 10, "fixed_schedule": True,
     "active": True, "flags": [], "role_preference": "manager",
     "default_availability": {
         "MON": "both", "TUE": "off",
         "WED": "night", "THU": "night",
         "FRI": "both", "SAT": "both", "SUN": "both",
     }},
    {"name": "Chan", "roles": ["manager"], "seniority": 10, "fixed_schedule": True,
     "active": True, "flags": [], "role_preference": "manager",
     "default_availability": {
         "MON": "off", "TUE": "both",
         "WED": "morning", "THU": "morning", "FRI": "morning",
         "SAT": "off", "SUN": "off",
     }},
]

DEFAULT_CONFIG = {
    "managers": {"lunch": "Aaron", "dinner": "Aaron"},
    "staffing": {
        day: {
            "lunch_servers": 3 if day in ("MON", "TUE", "WED", "THU") else 5,
            "lunch_hosts": 2,
            "mid_servers": 1 if day in ("MON", "TUE", "WED", "THU") else 2,
            "dinner_servers": 5 if day in ("MON", "TUE", "WED", "THU") else 8,
            "dinner_hosts": 3,
            "hibachi_lunch": 1,
            "hibachi_dinner": 2 if day in ("FRI", "SAT", "SUN") else 1,
        }
        for day in DAYS
    },
    "export_path": "",
    "week_start": "",
}


# ---------------------------------------------------------------------------
# Data Management
# ---------------------------------------------------------------------------
class DataManager:
    def __init__(self):
        DATA_DIR.mkdir(exist_ok=True)
        (DATA_DIR / "schedule_history").mkdir(exist_ok=True)
        self.staff_file = DATA_DIR / "staff.json"
        self.config_file = DATA_DIR / "config.json"
        self.availability_file = DATA_DIR / "availability.json"
        self.staff = self._load_staff()
        self.config = self._load_config()
        self.availability = self._load_availability()

    def _load_staff(self):
        if self.staff_file.exists():
            with open(self.staff_file) as f:
                return json.load(f)
        return copy.deepcopy(DEFAULT_STAFF)

    def _load_config(self):
        if self.config_file.exists():
            with open(self.config_file) as f:
                saved = json.load(f)
            # Merge with defaults for any missing keys
            merged = copy.deepcopy(DEFAULT_CONFIG)
            merged.update(saved)
            for day in DAYS:
                if day in saved.get("staffing", {}):
                    merged["staffing"][day].update(saved["staffing"][day])
            return merged
        return copy.deepcopy(DEFAULT_CONFIG)

    def _load_availability(self):
        if self.availability_file.exists():
            with open(self.availability_file) as f:
                return json.load(f)
        return {}

    def save_staff(self):
        with open(self.staff_file, "w") as f:
            json.dump(self.staff, f, indent=2)

    def save_config(self):
        with open(self.config_file, "w") as f:
            json.dump(self.config, f, indent=2)

    def save_availability(self):
        with open(self.availability_file, "w") as f:
            json.dump(self.availability, f, indent=2)

    def get_active_staff(self, role=None):
        result = [s for s in self.staff if s["active"]]
        if role:
            result = [s for s in result if role in s["roles"]]
        return result

    def get_staff_by_name(self, name):
        for s in self.staff:
            if s["name"] == name:
                return s
        return None


# ---------------------------------------------------------------------------
# Schedule Generator
# ---------------------------------------------------------------------------
class ScheduleGenerator:
    def __init__(self, data: DataManager):
        self.data = data
        self.schedule = {
            day: {
                "lunch_servers": [],
                "lunch_hosts": [],
                "lunch_manager": "",
                "mid_servers": [],
                "dinner_servers": [],
                "dinner_hosts": [],
                "dinner_manager": "",
                "hibachi_lunch_servers": [],
                "hibachi_servers": [],
            }
            for day in DAYS
        }
        self.shift_counts = {}  # name -> total shifts
        self.hibachi_counts = {}  # name -> hibachi shifts this week
        self.closing_counts = {}  # name -> closing host shifts this week
        self.warnings = []

    def generate(self):
        self.shift_counts = {}
        self.hibachi_counts = {}
        self.closing_counts = {}
        self.warnings = []
        avail = self.data.availability
        config = self.data.config

        # Reset schedule
        for day in DAYS:
            for key in self.schedule[day]:
                if isinstance(self.schedule[day][key], list):
                    self.schedule[day][key] = []
                else:
                    self.schedule[day][key] = ""

        # Step 1: Place managers based on their availability
        for day in DAYS:
            aaron_avail = self._get_availability("Aaron", day)
            chan_avail = self._get_availability("Chan", day)

            aaron_morning = aaron_avail in ("morning", "both")
            aaron_night = aaron_avail in ("night", "both")
            chan_morning = chan_avail in ("morning", "both")
            chan_night = chan_avail in ("night", "both")

            # Lunch manager(s)
            lunch_mgrs = []
            if aaron_morning:
                lunch_mgrs.append("Aaron")
            if chan_morning:
                lunch_mgrs.append("Chan")
            if not lunch_mgrs:
                lunch_mgrs.append(config["managers"].get("lunch", "Aaron"))
            self.schedule[day]["lunch_manager"] = "/".join(lunch_mgrs)

            # Dinner manager(s)
            dinner_mgrs = []
            if aaron_night:
                dinner_mgrs.append("Aaron")
            if chan_night:
                dinner_mgrs.append("Chan")
            if not dinner_mgrs:
                dinner_mgrs.append(config["managers"].get("dinner", "Aaron"))
            self.schedule[day]["dinner_manager"] = "/".join(dinner_mgrs)

        # Step 2: Place fixed schedule staff (Dian & Leony)
        self._place_fixed_staff(avail, config)

        # Step 3: Place Will (always hibachi)
        self._place_will(avail, config)

        # Step 4: Place hosts with seniority (Olivia first)
        self._place_hosts(avail, config)

        # Step 5: Place remaining servers
        self._place_servers(avail, config)

        # Step 6: Balance hibachi (ensure ~2 per server per week)
        self._balance_hibachi(config)

        # Step 7: Assign sidework letters (after hibachi so hibachi gets last letters)
        self._assign_sidework_letters()

        # Step 8: Check shortages (after all placement is done)
        self._check_shortages(config)

        return self.schedule, self.warnings

    def _get_availability(self, name, day):
        avail = self.data.availability
        if name in avail and day in avail[name]:
            return avail[name][day]
        return "off"

    def _can_work(self, name, day, shift):
        a = self._get_availability(name, day)
        if a == "both":
            return True
        if a == "morning" and shift in ("morning", "mid"):
            return True
        if a == "night" and shift in ("night", "mid"):
            return True
        return False

    def _is_assigned(self, name, day, shift=None):
        """Check if staff is already assigned. If shift is given, only check that shift."""
        s = self.schedule[day]
        lunch_keys = ["lunch_servers", "lunch_hosts", "hibachi_lunch_servers"]
        dinner_keys = ["dinner_servers", "dinner_hosts", "hibachi_servers"]
        if shift == "morning":
            keys = lunch_keys
        elif shift == "night":
            keys = dinner_keys
        else:
            keys = lunch_keys + ["mid_servers"] + dinner_keys
        for key in keys:
            if name in s[key]:
                return True
        return False

    def _add_shift(self, name):
        self.shift_counts[name] = self.shift_counts.get(name, 0) + 1

    def _get_shift_count(self, name):
        return self.shift_counts.get(name, 0)

    def _place_fixed_staff(self, avail, config):
        for name in ["Dian", "Leony"]:
            staff = self.data.get_staff_by_name(name)
            if not staff or not staff["active"]:
                continue
            for day in DAYS:
                a = self._get_availability(name, day)
                if a == "off":
                    continue
                staffing = config["staffing"][day]
                # Place in lunch if available for morning and lunch needs servers
                if a in ("morning", "both"):
                    if len(self.schedule[day]["lunch_servers"]) < staffing["lunch_servers"]:
                        self.schedule[day]["lunch_servers"].append(name)
                        self._add_shift(name)
                # Place in dinner if available for night
                if a in ("night", "both"):
                    if len(self.schedule[day]["dinner_servers"]) < staffing["dinner_servers"]:
                        self.schedule[day]["dinner_servers"].append(name)
                        self._add_shift(name)

    def _place_will(self, avail, config):
        staff = self.data.get_staff_by_name("Will")
        if not staff or not staff["active"]:
            return
        for day in DAYS:
            a = self._get_availability("Will", day)
            if a == "off":
                continue
            staffing = config["staffing"][day]
            if a in ("night", "both"):
                self.schedule[day]["hibachi_servers"].append("Will")
                self.hibachi_counts["Will"] = self.hibachi_counts.get("Will", 0) + 1
                self._add_shift("Will")
            elif a == "morning":
                if len(self.schedule[day]["lunch_servers"]) < staffing["lunch_servers"]:
                    self.schedule[day]["lunch_servers"].append("Will")
                    self._add_shift("Will")

    def _place_hosts(self, avail, config):
        hosts = [s for s in self.data.get_active_staff("host")
                 if not s["fixed_schedule"] and "emergency_only" not in s["flags"]]
        # Sort by seniority descending
        hosts.sort(key=lambda s: s["seniority"], reverse=True)

        for day in DAYS:
            staffing = config["staffing"][day]
            # Lunch hosts
            needed = staffing["lunch_hosts"]
            candidates = [h for h in hosts
                          if self._can_work(h["name"], day, "morning")
                          and not self._is_assigned(h["name"], day, "morning")]
            # Sort: fewer shifts first (balance), then seniority for ties
            candidates.sort(key=lambda h: (self._get_shift_count(h["name"]), -h["seniority"]))
            for h in candidates[:needed - len(self.schedule[day]["lunch_hosts"])]:
                self.schedule[day]["lunch_hosts"].append(h["name"])
                self._add_shift(h["name"])

            # Dinner hosts — balance closing (last slot)
            needed = staffing["dinner_hosts"]
            already = len(self.schedule[day]["dinner_hosts"])
            slots_to_fill = needed - already
            if slots_to_fill > 0:
                # Candidates who can close (excludes no_closing flag)
                can_close = [h for h in hosts
                             if self._can_work(h["name"], day, "night")
                             and not self._is_assigned(h["name"], day, "night")
                             and "no_closing" not in (self.data.get_staff_by_name(h["name"]) or {}).get("flags", [])]
                # Candidates who can't close (Maria etc.) — only for non-closing slots
                no_close = [h for h in hosts
                            if self._can_work(h["name"], day, "night")
                            and not self._is_assigned(h["name"], day, "night")
                            and "no_closing" in (self.data.get_staff_by_name(h["name"]) or {}).get("flags", [])]

                # Sort non-closing by shift balance
                no_close.sort(key=lambda h: (self._get_shift_count(h["name"]), -h["seniority"]))
                # Sort closing candidates by shift balance
                can_close.sort(key=lambda h: (self._get_shift_count(h["name"]), -h["seniority"]))

                # Fill non-closing slots first (all but last), prefer no_close staff for these
                non_closing_slots = slots_to_fill - 1
                closing_slot = 1
                assigned_names = set()

                # Pick non-closing slot hosts from both pools (no_close first, then can_close)
                non_closing_pool = no_close + can_close
                for h in non_closing_pool:
                    if non_closing_slots <= 0:
                        break
                    if h["name"] not in assigned_names:
                        self.schedule[day]["dinner_hosts"].append(h["name"])
                        self._add_shift(h["name"])
                        assigned_names.add(h["name"])
                        non_closing_slots -= 1

                # Pick closing slot host — fewest closing counts
                closing_candidates = [h for h in can_close if h["name"] not in assigned_names]
                closing_candidates.sort(key=lambda h: (
                    self.closing_counts.get(h["name"], 0),
                    self._get_shift_count(h["name"]),
                    -h["seniority"]))
                for h in closing_candidates[:closing_slot]:
                    self.schedule[day]["dinner_hosts"].append(h["name"])
                    self._add_shift(h["name"])
                    self.closing_counts[h["name"]] = self.closing_counts.get(h["name"], 0) + 1

        # Check if we need dual-role staff to fill host gaps
        for day in DAYS:
            staffing = config["staffing"][day]
            for shift_key, shift_type, needed_key in [
                ("lunch_hosts", "morning", "lunch_hosts"),
                ("dinner_hosts", "night", "dinner_hosts"),
            ]:
                needed = staffing[needed_key]
                current = len(self.schedule[day][shift_key])
                if current < needed:
                    # Try dual-role staff who prefer hosting
                    dual = [s for s in self.data.get_active_staff()
                            if "host" in s["roles"] and "server" in s["roles"]
                            and s["role_preference"] == "host"
                            and not s["fixed_schedule"]
                            and "emergency_only" not in s["flags"]
                            and self._can_work(s["name"], day, shift_type)
                            and not self._is_assigned(s["name"], day, shift_type)]
                    dual.sort(key=lambda s: self._get_shift_count(s["name"]))
                    for s in dual[:needed - current]:
                        self.schedule[day][shift_key].append(s["name"])
                        self._add_shift(s["name"])

    def _place_servers(self, avail, config):
        all_servers = [s for s in self.data.get_active_staff("server")
                       if not s["fixed_schedule"]
                       and "always_hibachi" not in s["flags"]
                       and "fill_in" not in s["flags"]
                       and "emergency_only" not in s["flags"]
                       and s.get("role_preference") != "host"]

        fill_ins = [s for s in self.data.get_active_staff("server")
                    if "fill_in" in s["flags"]]

        for day in DAYS:
            # Sort fill-ins: deprioritize those who prefer this day off
            def fill_in_sort_key(s):
                preferred_off = s.get("preferred_off", [])
                penalty = 10 if day in preferred_off else 0
                return (penalty, self._get_shift_count(s["name"]))

            staffing = config["staffing"][day]

            # Lunch servers (hibachi is picked from this pool, not additional)
            needed = staffing["lunch_servers"]
            current = len(self.schedule[day]["lunch_servers"])
            if current < needed:
                candidates = [s for s in all_servers
                              if self._can_work(s["name"], day, "morning")
                              and not self._is_assigned(s["name"], day, "morning")]
                # Deprioritize staff who prefer this day off
                def server_sort_key(s):
                    preferred_off = s.get("preferred_off", [])
                    off_penalty = 5 if day in preferred_off else 0
                    return (off_penalty, self._get_shift_count(s["name"]), -s["seniority"])
                candidates.sort(key=server_sort_key)
                for s in candidates[:needed - current]:
                    self.schedule[day]["lunch_servers"].append(s["name"])
                    self._add_shift(s["name"])

            # Still short? Try dual-role staff
            current = len(self.schedule[day]["lunch_servers"])
            if current < needed:
                dual = [s for s in self.data.get_active_staff()
                        if "server" in s["roles"] and s["role_preference"] == "server"
                        and not s["fixed_schedule"]
                        and "fill_in" not in s["flags"]
                        and "emergency_only" not in s["flags"]
                        and self._can_work(s["name"], day, "morning")
                        and not self._is_assigned(s["name"], day, "morning")]
                dual.sort(key=lambda s: self._get_shift_count(s["name"]))
                for s in dual[:needed - current]:
                    self.schedule[day]["lunch_servers"].append(s["name"])
                    self._add_shift(s["name"])

            # Still short? Try fill-ins
            current = len(self.schedule[day]["lunch_servers"])
            if current < needed:
                fi = [s for s in fill_ins
                      if self._can_work(s["name"], day, "morning")
                      and not self._is_assigned(s["name"], day, "morning")]
                fi.sort(key=fill_in_sort_key)
                for s in fi[:needed - current]:
                    self.schedule[day]["lunch_servers"].append(s["name"])
                    self._add_shift(s["name"])

            # Dinner servers (hibachi is picked from this pool, not additional)
            needed = staffing["dinner_servers"]
            current = len(self.schedule[day]["dinner_servers"])
            if current < needed:
                candidates = [s for s in all_servers
                              if self._can_work(s["name"], day, "night")
                              and not self._is_assigned(s["name"], day, "night")]
                def dinner_sort_key(s):
                    preferred_off = s.get("preferred_off", [])
                    off_penalty = 5 if day in preferred_off else 0
                    return (off_penalty, self._get_shift_count(s["name"]), -s["seniority"])
                candidates.sort(key=dinner_sort_key)
                for s in candidates[:needed - current]:
                    self.schedule[day]["dinner_servers"].append(s["name"])
                    self._add_shift(s["name"])

            # Dinner - try dual-role
            current = len(self.schedule[day]["dinner_servers"])
            if current < needed:
                dual = [s for s in self.data.get_active_staff()
                        if "server" in s["roles"] and s["role_preference"] == "server"
                        and not s["fixed_schedule"]
                        and "fill_in" not in s["flags"]
                        and "emergency_only" not in s["flags"]
                        and self._can_work(s["name"], day, "night")
                        and not self._is_assigned(s["name"], day, "night")]
                dual.sort(key=lambda s: self._get_shift_count(s["name"]))
                for s in dual[:needed - current]:
                    self.schedule[day]["dinner_servers"].append(s["name"])
                    self._add_shift(s["name"])

            # Dinner - try fill-ins
            current = len(self.schedule[day]["dinner_servers"])
            if current < needed:
                fi = [s for s in fill_ins
                      if self._can_work(s["name"], day, "night")
                      and not self._is_assigned(s["name"], day, "night")]
                fi.sort(key=fill_in_sort_key)
                for s in fi[:needed - current]:
                    self.schedule[day]["dinner_servers"].append(s["name"])
                    self._add_shift(s["name"])

            # Midshift
            needed_mid = staffing["mid_servers"]
            # Mid servers are selected by manager later, but we can suggest
            # For now leave empty - manager picks in the GUI


    def _check_shortages(self, config):
        """Check for staffing shortages after all placement is done."""
        for day in DAYS:
            staffing = config["staffing"][day]
            for key, label in [("lunch_servers", "Lunch Servers"),
                               ("dinner_servers", "Dinner Servers"),
                               ("lunch_hosts", "Lunch Hosts"),
                               ("dinner_hosts", "Dinner Hosts")]:
                if key == "lunch_hosts":
                    n = staffing["lunch_hosts"]
                elif key == "dinner_hosts":
                    n = staffing["dinner_hosts"]
                elif key == "lunch_servers":
                    n = staffing["lunch_servers"]
                else:
                    n = staffing["dinner_servers"]
                actual = len(self.schedule[day][key])
                # Include hibachi servers in the count
                if key == "dinner_servers":
                    actual += len(self.schedule[day]["hibachi_servers"])
                elif key == "lunch_servers":
                    actual += len(self.schedule[day]["hibachi_lunch_servers"])
                if actual < n:
                    self.warnings.append(
                        f"{day}: Need {n} {label} but only found {actual}")

    def _find_a_holder(self, candidates, priority_chain):
        """Find who should get letter A from a priority chain.
        Returns the first person in priority_chain who is in candidates,
        or None if nobody from the chain is available."""
        for name in priority_chain:
            if name in candidates:
                return name
        return None

    def _assign_sidework_letters(self):
        weekdays = ["MON", "TUE", "WED", "THU"]

        # Morning letter-A priority:
        #   Fri/Sat: Andy -> Dian/Leony -> Abigail -> Bryan -> random
        #   Weekdays: Dian -> Leony -> Abigail -> Bryan -> random
        #   Sun: Dian -> Leony -> Abigail -> Bryan -> random
        MORNING_A_PRIORITY_FRISATAM = ["Andy", "Dian", "Leony", "Abigail", "Bryan"]
        MORNING_A_PRIORITY_DEFAULT = ["Dian", "Leony", "Abigail", "Bryan"]

        # Night letter-A priority:
        #   Andy -> Dian -> Leony -> Sadie -> Bryan -> random
        NIGHT_A_PRIORITY = ["Andy", "Dian", "Leony", "Sadie", "Bryan"]

        for day in DAYS:
            # --- Morning servers get A-D ---
            # Build full letter list for total people, then split: regulars first, hibachi last
            servers = list(self.schedule[day]["lunch_servers"])
            hibachi_lunch = list(self.schedule[day]["hibachi_lunch_servers"])
            total = len(servers) + len(hibachi_lunch)

            # Build enough letters for everyone: A-D first, then doubles in order
            letters = list(MORNING_SIDEWORK)
            if total > len(letters):
                for extra_letter in MORNING_EXTRA_ORDER:
                    letters.append(extra_letter)
                    if len(letters) >= total:
                        break

            self.schedule[day]["lunch_server_letters"] = {}

            # Determine morning A holder (from regular servers only)
            if day in ("FRI", "SAT"):
                morning_a = self._find_a_holder(servers, MORNING_A_PRIORITY_FRISATAM)
            else:
                morning_a = self._find_a_holder(servers, MORNING_A_PRIORITY_DEFAULT)

            if morning_a:
                ordered_regular = [morning_a] + [n for n in servers if n != morning_a]
            else:
                ordered_regular = servers

            # Regular servers get the first letters, hibachi get the last
            for i, name in enumerate(ordered_regular):
                self.schedule[day]["lunch_server_letters"][name] = letters[i]

            for i, name in enumerate(hibachi_lunch):
                idx = len(ordered_regular) + i
                self.schedule[day]["lunch_server_letters"][name] = letters[idx] if idx < len(letters) else letters[-1]

            # --- Night servers get A-H ---
            # Regular servers first, hibachi last (hibachi gets last letters)
            dinner_regular = list(self.schedule[day]["dinner_servers"])
            dinner_hibachi = list(self.schedule[day]["hibachi_servers"])
            all_night = dinner_regular + dinner_hibachi
            letters = list(NIGHT_SIDEWORK)
            if len(all_night) > len(letters):
                for extra_letter in NIGHT_EXTRA_ORDER:
                    letters.append(extra_letter)
                    if len(letters) >= len(all_night):
                        break

            self.schedule[day]["dinner_server_letters"] = {}

            night_a = self._find_a_holder(dinner_regular, NIGHT_A_PRIORITY)

            if night_a:
                ordered_night = [night_a] + [n for n in dinner_regular if n != night_a]
            else:
                ordered_night = dinner_regular

            # Build full letter list for total count
            total_night = len(ordered_night) + len(dinner_hibachi)
            if total_night > len(letters):
                for extra_letter in NIGHT_EXTRA_ORDER:
                    letters.append(extra_letter)
                    if len(letters) >= total_night:
                        break

            # Regular servers get the first letters, hibachi get the last
            for i, name in enumerate(ordered_night):
                self.schedule[day]["dinner_server_letters"][name] = letters[i]

            for i, name in enumerate(dinner_hibachi):
                idx = len(ordered_night) + i
                self.schedule[day]["dinner_server_letters"][name] = letters[idx] if idx < len(letters) else letters[-1]


    def _balance_hibachi(self, config):
        # Each server should get ~2 hibachi shifts per week
        # Will is excluded (always hibachi)
        all_servers = [s["name"] for s in self.data.get_active_staff("server")
                       if "always_hibachi" not in s["flags"]
                       and "emergency_only" not in s["flags"]
                       and s["active"]]

        # Track who has hibachi assignments
        for name in all_servers:
            self.hibachi_counts[name] = 0

        # Assign lunch hibachi — pick from lunch_servers pool
        for day in DAYS:
            staffing = config["staffing"][day]
            needed = staffing.get("hibachi_lunch", 0)
            already = len(self.schedule[day]["hibachi_lunch_servers"])
            needed = needed - already
            if needed <= 0:
                continue
            lunch = self.schedule[day]["lunch_servers"][:]
            lunch_eligible = [n for n in lunch if n in all_servers
                              and self.hibachi_counts.get(n, 0) < 2]
            # Sort: fewest hibachi first, then lowest seniority first (so senior staff stay on regular sidework)
            lunch_eligible.sort(key=lambda n: (
                self.hibachi_counts.get(n, 0),
                -(self.data.get_staff_by_name(n) or {}).get("seniority", 0)))

            picked = lunch_eligible[:needed]
            # Fallback: if not enough with < 2 count, allow anyone
            if len(picked) < needed:
                remaining = needed - len(picked)
                picked_names = set(n for n in picked)
                fallback = [n for n in lunch if n in all_servers and n not in picked_names]
                fallback.sort(key=lambda n: self.hibachi_counts.get(n, 0))
                picked.extend(fallback[:remaining])
            for name in picked:
                self.hibachi_counts[name] = self.hibachi_counts.get(name, 0) + 1
                if name in self.schedule[day]["lunch_servers"]:
                    self.schedule[day]["lunch_servers"].remove(name)
                self.schedule[day]["hibachi_lunch_servers"].append(name)

        # Assign dinner hibachi — pick from dinner_servers pool
        for day in DAYS:
            staffing = config["staffing"][day]
            needed = staffing.get("hibachi_dinner", 0)
            already = len(self.schedule[day]["hibachi_servers"])
            needed = needed - already
            if needed <= 0:
                continue
            dinner = self.schedule[day]["dinner_servers"][:]
            dinner_eligible = [n for n in dinner if n in all_servers
                               and self.hibachi_counts.get(n, 0) < 2]
            # Sort: fewest hibachi first, then lowest seniority first
            dinner_eligible.sort(key=lambda n: (
                self.hibachi_counts.get(n, 0),
                -(self.data.get_staff_by_name(n) or {}).get("seniority", 0)))

            picked = dinner_eligible[:needed]
            # Fallback: if not enough with < 2 count, allow anyone
            if len(picked) < needed:
                remaining = needed - len(picked)
                picked_names = set(n for n in picked)
                fallback = [n for n in dinner if n in all_servers and n not in picked_names]
                fallback.sort(key=lambda n: self.hibachi_counts.get(n, 0))
                picked.extend(fallback[:remaining])
            for name in picked:
                self.hibachi_counts[name] = self.hibachi_counts.get(name, 0) + 1
                if name in self.schedule[day]["dinner_servers"]:
                    self.schedule[day]["dinner_servers"].remove(name)
                self.schedule[day]["hibachi_servers"].append(name)

        # Second pass: for servers who still have 0 hibachi, try to swap them
        # into a hibachi slot by replacing someone who already has 2
        for name in all_servers:
            if self.hibachi_counts.get(name, 0) >= 1:
                continue
            # Find a day this server works dinner where we can swap
            for day in DAYS:
                if name not in self.schedule[day]["dinner_servers"]:
                    continue
                # Find someone in hibachi on this day with count >= 2
                for hib_name in list(self.schedule[day]["hibachi_servers"]):
                    if hib_name == "Will":
                        continue
                    if self.hibachi_counts.get(hib_name, 0) >= 2:
                        # Swap: move hib_name back to dinner, move name to hibachi
                        self.schedule[day]["hibachi_servers"].remove(hib_name)
                        self.schedule[day]["dinner_servers"].append(hib_name)
                        self.schedule[day]["dinner_servers"].remove(name)
                        self.schedule[day]["hibachi_servers"].append(name)
                        self.hibachi_counts[hib_name] -= 1
                        self.hibachi_counts[name] = self.hibachi_counts.get(name, 0) + 1
                        break
                if self.hibachi_counts.get(name, 0) >= 1:
                    break

        # Check balance and warn
        for name in all_servers:
            count = self.hibachi_counts.get(name, 0)
            if count == 0:
                self.warnings.append(f"{name} has 0 hibachi shifts (target: 2)")
            elif count > 2:
                self.warnings.append(f"{name} has {count} hibachi shifts (target: 2)")


# ---------------------------------------------------------------------------
# Excel Exporter
# ---------------------------------------------------------------------------
class ExcelExporter:
    def __init__(self, schedule, config, week_start_date=None):
        self.schedule = schedule
        self.config = config
        self.week_start = week_start_date

    def export(self, filepath):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Chart2"

        thin = Side(style="thin")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        header_font = Font(bold=True, size=11)
        name_font = Font(size=10)
        time_font = Font(size=9, italic=True)
        mid_font = Font(size=10, color="0000FF")

        # Column widths
        ws.column_dimensions["A"].width = 16
        for col in ["B", "C", "D", "E", "F", "G", "H"]:
            ws.column_dimensions[col].width = 18

        # Row 1: Dates
        if self.week_start:
            for i, col in enumerate(["B", "C", "D", "E", "F", "G", "H"]):
                d = self.week_start + timedelta(days=i)
                ws[f"{col}1"] = d.strftime("%m/%d/%Y")
                ws[f"{col}1"].font = header_font
                ws[f"{col}1"].alignment = Alignment(horizontal="center")

        # Row 2: Day names
        for i, (col, day) in enumerate(zip(["B", "C", "D", "E", "F", "G", "H"], DAYS)):
            ws[f"{col}2"] = day
            ws[f"{col}2"].font = header_font
            ws[f"{col}2"].alignment = Alignment(horizontal="center")

        # LUNCH SERVERS (rows 3-8)
        ws["A4"] = "Servers"
        ws["A4"].font = header_font
        ws["A6"] = LUNCH_SERVER_TIME
        ws["A6"].font = time_font

        for di, day in enumerate(DAYS):
            col = chr(66 + di)  # B-H
            hibachi_lunch = self.schedule[day]["hibachi_lunch_servers"]
            servers = self.schedule[day]["lunch_servers"]
            letters = self.schedule[day].get("lunch_server_letters", {})
            mids = self.schedule[day].get("mid_servers", [])

            # Place lunch hibachi first (gold background)
            row = 3
            for name in hibachi_lunch:
                letter = letters.get(name, chr(65 + row - 3))
                mid_mark = "*" if name in mids else ""
                ws[f"{col}{row}"] = f"{name} {letter}{mid_mark}"
                ws[f"{col}{row}"].font = name_font
                ws[f"{col}{row}"].alignment = Alignment(horizontal="center")
                ws[f"{col}{row}"].fill = PatternFill(start_color="FFE4B5", end_color="FFE4B5", fill_type="solid")
                row += 1

            # Then regular lunch servers
            for name in servers:
                letter = letters.get(name, chr(65 + row - 3))
                mid_mark = "*" if name in mids else ""
                ws[f"{col}{row}"] = f"{name} {letter}{mid_mark}"
                ws[f"{col}{row}"].font = name_font
                ws[f"{col}{row}"].alignment = Alignment(horizontal="center")
                row += 1

        # LUNCH MANAGER (rows 13-14)
        ws["A13"] = "Lunch"
        ws["A13"].font = header_font
        ws["A14"] = "Manager"
        ws["A14"].font = header_font
        for di, day in enumerate(DAYS):
            col = chr(66 + di)
            ws[f"{col}13"] = self.schedule[day]["lunch_manager"]
            ws[f"{col}13"].font = name_font
            ws[f"{col}13"].alignment = Alignment(horizontal="center")
            ws[f"{col}14"] = LUNCH_MANAGER_TIME
            ws[f"{col}14"].font = time_font
            ws[f"{col}14"].alignment = Alignment(horizontal="center")

        # LUNCH HOSTS (rows 15-19)
        ws["A15"] = "Hostess"
        ws["A15"].font = header_font
        ws["A17"] = "Hostess"
        ws["A17"].font = header_font
        ws["A19"] = "Hostess"
        ws["A19"].font = header_font

        for di, day in enumerate(DAYS):
            col = chr(66 + di)
            hosts = self.schedule[day]["lunch_hosts"]
            for hi, name in enumerate(hosts):
                if hi == 0:
                    ws[f"{col}15"] = name
                    ws[f"{col}16"] = LUNCH_HOST_TIMES[0]
                elif hi == 1:
                    ws[f"{col}17"] = name
                    ws[f"{col}18"] = LUNCH_HOST_TIMES[1]
                for r in [15, 16, 17, 18]:
                    if ws[f"{col}{r}"].value:
                        ws[f"{col}{r}"].font = name_font
                        ws[f"{col}{r}"].alignment = Alignment(horizontal="center")

        # Transition time row 20
        time_row_20 = {
            "MON": "(10:15-2:15)", "TUE": "(10:15-4:15)",
            "WED": "(4:00-8:00)", "THU": "(10:15-2:15)",
            "FRI": "(10:15-3:15)", "SAT": "(10:15-2:15)",
            "SUN": "(11:30-3:15)",
        }
        for di, day in enumerate(DAYS):
            col = chr(66 + di)
            ws[f"{col}20"] = time_row_20.get(day, "")
            ws[f"{col}20"].font = time_font
            ws[f"{col}20"].alignment = Alignment(horizontal="center")

        # DINNER SERVERS (rows 21-29)
        ws["A22"] = "Servers"
        ws["A22"].font = header_font
        ws["A23"] = "(s)"
        ws["A23"].font = time_font
        ws["A27"] = DINNER_SERVER_TIME
        ws["A27"].font = time_font

        for di, day in enumerate(DAYS):
            col = chr(66 + di)
            servers = self.schedule[day]["dinner_servers"]
            hibachi = self.schedule[day]["hibachi_servers"]
            letters = self.schedule[day].get("dinner_server_letters", {})
            mids = self.schedule[day].get("mid_servers", [])

            # Place hibachi first (gold background)
            row = 21
            for name in hibachi:
                letter = letters.get(name, "W")
                mid_mark = "*" if name in mids else ""
                ws[f"{col}{row}"] = f"{name} {letter}{mid_mark}"
                ws[f"{col}{row}"].font = name_font
                ws[f"{col}{row}"].alignment = Alignment(horizontal="center")
                ws[f"{col}{row}"].fill = PatternFill(start_color="FFE4B5", end_color="FFE4B5", fill_type="solid")
                row += 1

            # Then regular dinner servers
            for name in servers:
                letter = letters.get(name, chr(65 + row - 21))
                mid_mark = "*" if name in mids else ""
                ws[f"{col}{row}"] = f"{name} {letter}{mid_mark}"
                ws[f"{col}{row}"].font = name_font
                ws[f"{col}{row}"].alignment = Alignment(horizontal="center")
                row += 1

        # DINNER HOSTS (rows 33-38)
        ws["A33"] = "Hostesses"
        ws["A33"].font = header_font
        ws["A34"] = "(H)"
        ws["A34"].font = time_font

        for di, day in enumerate(DAYS):
            col = chr(66 + di)
            hosts = self.schedule[day]["dinner_hosts"]
            day_idx = di  # 0=MON ... 6=SUN

            if day_idx >= 5:  # SAT, SUN
                times = DINNER_HOST_TIMES_WEEKEND if day != "SUN" else DINNER_HOST_TIMES_SUNDAY
            elif day_idx == 4:  # FRI
                times = DINNER_HOST_TIMES_FRIDAY
            else:
                times = DINNER_HOST_TIMES_WEEKDAY

            for hi, name in enumerate(hosts):
                name_row = 33 + hi * 2
                time_row = 34 + hi * 2
                ws[f"{col}{name_row}"] = name
                ws[f"{col}{name_row}"].font = name_font
                ws[f"{col}{name_row}"].alignment = Alignment(horizontal="center")
                if hi < len(times):
                    ws[f"{col}{time_row}"] = times[hi]
                    ws[f"{col}{time_row}"].font = time_font
                    ws[f"{col}{time_row}"].alignment = Alignment(horizontal="center")

        # DINNER MANAGER (rows 39-40)
        ws["A39"] = "Dinner"
        ws["A39"].font = header_font
        ws["A40"] = "Manager"
        ws["A40"].font = header_font
        for di, day in enumerate(DAYS):
            col = chr(66 + di)
            ws[f"{col}39"] = self.schedule[day]["dinner_manager"]
            ws[f"{col}39"].font = name_font
            ws[f"{col}39"].alignment = Alignment(horizontal="center")
            t = DINNER_MANAGER_TIME_WEEKEND if di >= 4 else DINNER_MANAGER_TIME_WEEKDAY
            ws[f"{col}40"] = t
            ws[f"{col}40"].font = time_font
            ws[f"{col}40"].alignment = Alignment(horizontal="center")

        # Legend
        ws["B41"] = "Gold = Hibachi Servers"
        ws["B41"].font = Font(bold=True, size=9)
        ws["A42"] = "*"
        ws["A42"].font = Font(bold=True, size=11)
        ws["B42"] = f"Mid shift server(s)  {MIDSHIFT_TIME}"
        ws["B42"].font = Font(size=9, italic=True)

        wb.save(filepath)
        return filepath


# ---------------------------------------------------------------------------
# PDF Exporter
# ---------------------------------------------------------------------------
class PDFExporter:
    def __init__(self, schedule, config, week_start_date=None):
        self.schedule = schedule
        self.config = config
        self.week_start = week_start_date

    def export(self, filepath):
        if FPDF is None:
            raise RuntimeError("fpdf2 not installed. Run setup.sh.")

        pdf = FPDF(orientation="L", unit="mm", format="A4")
        pdf.add_page()
        pdf.set_auto_page_break(auto=False)

        col_w = [30] + [35] * 7  # A column + 7 day columns
        x_start = 5
        y_start = 10

        def draw_row(y, texts, bold=False, size=7):
            pdf.set_font("Helvetica", "B" if bold else "", size)
            for i, txt in enumerate(texts):
                x = x_start + sum(col_w[:i])
                pdf.set_xy(x, y)
                pdf.cell(col_w[i], 5, str(txt), border=1, align="C")
            return y + 5

        # Title
        title = "Toyo Schedule"
        if self.week_start:
            end = self.week_start + timedelta(days=6)
            title += f"  {self.week_start.strftime('%b %d')} - {end.strftime('%b %d, %Y')}"
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_xy(x_start, y_start)
        pdf.cell(sum(col_w), 7, title, align="C")
        y = y_start + 9

        # Day headers
        headers = [""] + DAYS
        if self.week_start:
            headers = [""] + [(self.week_start + timedelta(days=i)).strftime("%a %m/%d") for i in range(7)]
        y = draw_row(y, headers, bold=True, size=8)

        # LUNCH SECTION
        y = draw_row(y, ["LUNCH SERVERS", "", "", "", "", "", "", ""], bold=True, size=7)

        # Find max lunch servers across days (including hibachi lunch)
        max_lunch = max(len(self.schedule[d]["lunch_servers"]) + len(self.schedule[d]["hibachi_lunch_servers"]) for d in DAYS) if DAYS else 0
        for si in range(max(max_lunch, 1)):
            row = [""]
            if si == 0:
                row[0] = LUNCH_SERVER_TIME
            for day in DAYS:
                hibachi_lunch = self.schedule[day]["hibachi_lunch_servers"]
                servers = self.schedule[day]["lunch_servers"]
                letters = self.schedule[day].get("lunch_server_letters", {})
                hib_set = set(hibachi_lunch)
                combined = servers + hibachi_lunch
                combined.sort(key=lambda n: (letters.get(n, "Z"), 1 if n in hib_set else 0))
                mids = self.schedule[day].get("mid_servers", [])
                if si < len(combined):
                    name = combined[si]
                    letter = letters.get(name, chr(65 + si))
                    mid = "*" if name in mids else ""
                    row.append(f"{name} {letter}{mid}")
                else:
                    row.append("")
            y = draw_row(y, row, size=6)

        # Lunch manager
        row_mgr = ["Lunch Manager"]
        for day in DAYS:
            row_mgr.append(self.schedule[day]["lunch_manager"])
        y = draw_row(y, row_mgr, bold=True, size=6)
        # Lunch manager time
        row_mgr_t = [""] + [LUNCH_MANAGER_TIME] * 7
        y = draw_row(y, row_mgr_t, size=5)

        # Lunch hosts — name row then time row for each hostess slot
        max_lh = max(len(self.schedule[d]["lunch_hosts"]) for d in DAYS) if DAYS else 0
        for hi in range(max(max_lh, 1)):
            # Name row
            row = [f"Hostess {hi+1}" if hi < 3 else ""]
            for day in DAYS:
                hosts = self.schedule[day]["lunch_hosts"]
                row.append(hosts[hi] if hi < len(hosts) else "")
            y = draw_row(y, row, size=6)
            # Time row
            row_t = [""]
            for day in DAYS:
                hosts = self.schedule[day]["lunch_hosts"]
                if hi < len(hosts) and hi < len(LUNCH_HOST_TIMES):
                    row_t.append(LUNCH_HOST_TIMES[hi])
                else:
                    row_t.append("")
            y = draw_row(y, row_t, size=5)

        # Transition time row
        row_trans = [""]
        time_row_20 = {
            "MON": "(10:15-2:15)", "TUE": "(10:15-4:15)",
            "WED": "(4:00-8:00)", "THU": "(10:15-2:15)",
            "FRI": "(10:15-3:15)", "SAT": "(10:15-2:15)",
            "SUN": "(11:30-3:15)",
        }
        for day in DAYS:
            row_trans.append(time_row_20.get(day, ""))
        y = draw_row(y, row_trans, size=5)

        y += 1  # spacer

        # DINNER SECTION
        y = draw_row(y, ["DINNER SERVERS", "", "", "", "", "", "", ""], bold=True, size=7)

        max_dinner = max(len(self.schedule[d]["dinner_servers"]) + len(self.schedule[d]["hibachi_servers"]) for d in DAYS)
        for si in range(max(max_dinner, 1)):
            row = [""]
            if si == 0:
                row[0] = DINNER_SERVER_TIME
            for day in DAYS:
                hibachi = self.schedule[day]["hibachi_servers"]
                servers = self.schedule[day]["dinner_servers"]
                letters = self.schedule[day].get("dinner_server_letters", {})
                hib_set = set(hibachi)
                combined = servers + hibachi
                combined.sort(key=lambda n: (letters.get(n, "Z"), 1 if n in hib_set else 0))
                mids = self.schedule[day].get("mid_servers", [])
                if si < len(combined):
                    name = combined[si]
                    letter = letters.get(name, chr(65 + si))
                    mid = "*" if name in mids else ""
                    row.append(f"{name} {letter}{mid}")
                else:
                    row.append("")
            y = draw_row(y, row, size=6)

        # Dinner hosts — name row then time row for each hostess slot
        y += 1
        max_dh = max(len(self.schedule[d]["dinner_hosts"]) for d in DAYS) if DAYS else 0
        for hi in range(max(max_dh, 1)):
            # Name row
            row = [f"Hostess {hi+1}" if hi < 3 else ""]
            for day in DAYS:
                hosts = self.schedule[day]["dinner_hosts"]
                row.append(hosts[hi] if hi < len(hosts) else "")
            y = draw_row(y, row, size=6)
            # Time row
            row_t = [""]
            for di, day in enumerate(DAYS):
                hosts = self.schedule[day]["dinner_hosts"]
                if hi < len(hosts):
                    if di >= 5:  # SAT, SUN
                        times = DINNER_HOST_TIMES_WEEKEND if day != "SUN" else DINNER_HOST_TIMES_SUNDAY
                    elif di == 4:  # FRI
                        times = DINNER_HOST_TIMES_FRIDAY
                    else:
                        times = DINNER_HOST_TIMES_WEEKDAY
                    row_t.append(times[hi] if hi < len(times) else "")
                else:
                    row_t.append("")
            y = draw_row(y, row_t, size=5)

        # Dinner manager
        row_mgr = ["Dinner Manager"]
        for day in DAYS:
            row_mgr.append(self.schedule[day]["dinner_manager"])
        y = draw_row(y, row_mgr, bold=True, size=6)
        # Dinner manager time
        row_mgr_t = [""]
        for di, day in enumerate(DAYS):
            t = DINNER_MANAGER_TIME_WEEKEND if di >= 4 else DINNER_MANAGER_TIME_WEEKDAY
            row_mgr_t.append(t)
        y = draw_row(y, row_mgr_t, size=5)

        # Legend
        y += 3
        pdf.set_font("Helvetica", "I", 7)
        pdf.set_xy(x_start, y)
        pdf.cell(0, 4, "* = Midshift server(s)  |  Gold = Hibachi side  |  Letters = Sidework assignment")

        pdf.output(filepath)
        return filepath


# ---------------------------------------------------------------------------
# Main Application GUI
# ---------------------------------------------------------------------------
class ToyoSchedulerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Toyo Schedule Maker")
        self.root.geometry("1200x750")
        self.root.minsize(1000, 600)

        self.data = DataManager()
        self.generator = ScheduleGenerator(self.data)
        self.current_schedule = None
        self.current_warnings = []

        style = ttk.Style()
        style.configure("TNotebook.Tab", padding=[12, 6], font=("Helvetica", 10))
        style.configure("Treeview", rowheight=24, font=("Helvetica", 9))
        style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"))

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self._build_staff_tab()
        self._build_config_tab()
        self._build_availability_tab()
        self._build_schedule_tab()
        self._build_export_tab()

    # ---- TAB 1: STAFF MANAGEMENT ----
    def _build_staff_tab(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="  Staff  ")

        # Toolbar
        toolbar = ttk.Frame(frame)
        toolbar.pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(toolbar, text="Add Staff", command=self._add_staff).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Edit Selected", command=self._edit_staff).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Remove Selected", command=self._remove_staff).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Reset to Defaults", command=self._reset_staff).pack(side=tk.RIGHT, padx=2)

        # Treeview
        cols = ("Name", "Roles", "Seniority", "Preference", "Flags", "Active")
        self.staff_tree = ttk.Treeview(frame, columns=cols, show="headings", height=20)
        for col in cols:
            self.staff_tree.heading(col, text=col)
            w = 80 if col != "Name" else 120
            if col == "Flags":
                w = 180
            self.staff_tree.column(col, width=w, anchor="center")

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.staff_tree.yview)
        self.staff_tree.configure(yscrollcommand=scrollbar.set)
        self.staff_tree.pack(fill=tk.BOTH, expand=True, padx=5, side=tk.LEFT)
        scrollbar.pack(fill=tk.Y, side=tk.LEFT)

        self.staff_tree.bind("<Double-1>", lambda e: self._edit_staff())
        self._refresh_staff_tree()

    def _refresh_staff_tree(self):
        for item in self.staff_tree.get_children():
            self.staff_tree.delete(item)
        for s in self.data.staff:
            roles = ", ".join(s["roles"])
            flags = ", ".join(s.get("flags", []))
            pref = s.get("role_preference", "") or ""
            active = "Yes" if s["active"] else "No"
            self.staff_tree.insert("", "end", values=(
                s["name"], roles, s["seniority"], pref, flags, active))

    def _add_staff(self):
        self._staff_dialog(None)

    def _edit_staff(self):
        sel = self.staff_tree.selection()
        if not sel:
            messagebox.showinfo("Select", "Select a staff member to edit.")
            return
        values = self.staff_tree.item(sel[0])["values"]
        name = values[0]
        staff = self.data.get_staff_by_name(name)
        if staff:
            self._staff_dialog(staff)

    def _remove_staff(self):
        sel = self.staff_tree.selection()
        if not sel:
            return
        values = self.staff_tree.item(sel[0])["values"]
        name = values[0]
        if messagebox.askyesno("Remove", f"Remove {name} from roster?"):
            self.data.staff = [s for s in self.data.staff if s["name"] != name]
            self.data.save_staff()
            self._refresh_staff_tree()

    def _reset_staff(self):
        if messagebox.askyesno("Reset", "Reset staff to factory defaults?"):
            self.data.staff = copy.deepcopy(DEFAULT_STAFF)
            self.data.save_staff()
            self._refresh_staff_tree()

    def _staff_dialog(self, existing):
        dlg = tk.Toplevel(self.root)
        dlg.title("Edit Staff" if existing else "Add Staff")
        dlg.geometry("400x450")
        dlg.transient(self.root)
        dlg.grab_set()

        row = 0
        ttk.Label(dlg, text="Name:").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        name_var = tk.StringVar(value=existing["name"] if existing else "")
        ttk.Entry(dlg, textvariable=name_var, width=25).grid(row=row, column=1, padx=10, pady=5)

        row += 1
        ttk.Label(dlg, text="Roles:").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        roles_frame = ttk.Frame(dlg)
        roles_frame.grid(row=row, column=1, padx=10, pady=5, sticky="w")
        server_var = tk.BooleanVar(value="server" in existing["roles"] if existing else False)
        host_var = tk.BooleanVar(value="host" in existing["roles"] if existing else False)
        ttk.Checkbutton(roles_frame, text="Server", variable=server_var).pack(side=tk.LEFT)
        ttk.Checkbutton(roles_frame, text="Host", variable=host_var).pack(side=tk.LEFT)

        row += 1
        ttk.Label(dlg, text="Seniority (1-10):").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        sen_var = tk.IntVar(value=existing["seniority"] if existing else 5)
        ttk.Spinbox(dlg, from_=1, to=10, textvariable=sen_var, width=5).grid(row=row, column=1, padx=10, pady=5, sticky="w")

        row += 1
        ttk.Label(dlg, text="Role Preference:").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        pref_var = tk.StringVar(value=existing.get("role_preference", "") or "none" if existing else "none")
        pref_combo = ttk.Combobox(dlg, textvariable=pref_var,
                                  values=["none", "server", "host"], state="readonly", width=10)
        pref_combo.grid(row=row, column=1, padx=10, pady=5, sticky="w")

        row += 1
        ttk.Label(dlg, text="Flags:").grid(row=row, column=0, padx=10, pady=5, sticky="nw")
        flags_frame = ttk.Frame(dlg)
        flags_frame.grid(row=row, column=1, padx=10, pady=5, sticky="w")
        existing_flags = existing.get("flags", []) if existing else []
        flag_vars = {}
        for flag in ["fixed_schedule", "always_hibachi", "no_closing",
                      "fill_in", "emergency_only", "seniority_priority"]:
            var = tk.BooleanVar(value=flag in existing_flags)
            ttk.Checkbutton(flags_frame, text=flag, variable=var).pack(anchor="w")
            flag_vars[flag] = var

        row += 1
        ttk.Label(dlg, text="Fixed Schedule:").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        fixed_var = tk.BooleanVar(value=existing.get("fixed_schedule", False) if existing else False)
        ttk.Checkbutton(dlg, variable=fixed_var).grid(row=row, column=1, padx=10, pady=5, sticky="w")

        row += 1
        ttk.Label(dlg, text="Active:").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        active_var = tk.BooleanVar(value=existing["active"] if existing else True)
        ttk.Checkbutton(dlg, variable=active_var).grid(row=row, column=1, padx=10, pady=5, sticky="w")

        def save():
            name = name_var.get().strip()
            if not name:
                messagebox.showerror("Error", "Name is required.", parent=dlg)
                return
            roles = []
            if server_var.get():
                roles.append("server")
            if host_var.get():
                roles.append("host")
            if not roles:
                messagebox.showerror("Error", "Select at least one role.", parent=dlg)
                return

            flags = [f for f, v in flag_vars.items() if v.get()]
            pref = pref_var.get()
            if pref == "none":
                pref = None

            new_staff = {
                "name": name,
                "roles": roles,
                "seniority": sen_var.get(),
                "fixed_schedule": fixed_var.get(),
                "active": active_var.get(),
                "flags": flags,
                "role_preference": pref,
            }

            if existing:
                # Update in place
                for i, s in enumerate(self.data.staff):
                    if s["name"] == existing["name"]:
                        self.data.staff[i] = new_staff
                        break
            else:
                # Check duplicate
                if self.data.get_staff_by_name(name):
                    messagebox.showerror("Error", f"{name} already exists.", parent=dlg)
                    return
                self.data.staff.append(new_staff)

            self.data.save_staff()
            self._refresh_staff_tree()
            dlg.destroy()

        row += 1
        ttk.Button(dlg, text="Save", command=save).grid(row=row, column=0, columnspan=2, pady=15)

    # ---- TAB 2: SHIFT CONFIGURATION ----
    def _build_config_tab(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="  Configuration  ")

        # Staffing grid
        grid_frame = ttk.LabelFrame(frame, text="Staffing Numbers Per Day")
        grid_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        labels = ["", "Lunch\nServers", "Lunch\nHosts", "Hibachi\nLunch",
                  "Dinner\nServers", "Dinner\nHosts", "Hibachi\nDinner"]
        for i, label in enumerate(labels):
            ttk.Label(grid_frame, text=label, font=("Helvetica", 9, "bold"),
                     justify="center").grid(row=0, column=i, padx=8, pady=5)

        self.config_vars = {}
        keys = ["lunch_servers", "lunch_hosts", "hibachi_lunch",
                "dinner_servers", "dinner_hosts", "hibachi_dinner"]

        for di, day in enumerate(DAYS):
            ttk.Label(grid_frame, text=day, font=("Helvetica", 10, "bold")).grid(
                row=di + 1, column=0, padx=10, pady=3)
            self.config_vars[day] = {}
            for ki, key in enumerate(keys):
                val = self.data.config["staffing"][day].get(key, 0)
                var = tk.IntVar(value=val)
                self.config_vars[day][key] = var
                sb = ttk.Spinbox(grid_frame, from_=0, to=15, textvariable=var,
                                width=4, justify="center")
                sb.grid(row=di + 1, column=ki + 1, padx=8, pady=3)

        # Save button
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        ttk.Button(btn_frame, text="Save Configuration", command=self._save_config).pack(side=tk.RIGHT)
        ttk.Button(btn_frame, text="Copy Weekday to All", command=self._copy_weekday).pack(side=tk.LEFT, padx=5)

    def _save_config(self):
        keys = ["lunch_servers", "lunch_hosts", "hibachi_lunch",
                "dinner_servers", "dinner_hosts", "hibachi_dinner"]
        for day in DAYS:
            for key in keys:
                self.data.config["staffing"][day][key] = self.config_vars[day][key].get()
        self.data.save_config()
        messagebox.showinfo("Saved", "Configuration saved!")

    def _copy_weekday(self):
        """Copy Monday values to Tue-Thu"""
        keys = ["lunch_servers", "lunch_hosts", "hibachi_lunch",
                "dinner_servers", "dinner_hosts", "hibachi_dinner"]
        for day in ["TUE", "WED", "THU"]:
            for key in keys:
                self.config_vars[day][key].set(self.config_vars["MON"][key].get())

    # ---- TAB 3: AVAILABILITY ----
    def _build_availability_tab(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="  Availability  ")

        # Toolbar
        toolbar = ttk.Frame(frame)
        toolbar.pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(toolbar, text="Set All Available", command=self._set_all_available).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Clear All", command=self._clear_availability).pack(side=tk.LEFT, padx=2)
        if HAS_OCR:
            ttk.Button(toolbar, text="Import from Photo", command=self._import_from_photo).pack(side=tk.LEFT, padx=8)
        ttk.Button(toolbar, text="Save Availability", command=self._save_availability).pack(side=tk.RIGHT, padx=2)

        # Scrollable grid
        canvas = tk.Canvas(frame)
        v_scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=canvas.yview)
        self.avail_frame = ttk.Frame(canvas)

        self.avail_frame.bind("<Configure>",
                              lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.avail_frame, anchor="nw")
        canvas.configure(yscrollcommand=v_scroll.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Bind mousewheel
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

        self._build_availability_grid()

    def _build_availability_grid(self):
        for widget in self.avail_frame.winfo_children():
            widget.destroy()

        # Internal avail_vars still tracks name -> day -> StringVar("off"/"morning"/"night"/"both")
        # for compatibility with save/load/photo import
        self.avail_vars = {}
        active = [s for s in self.data.staff if s["active"]]

        # Initialize avail_vars for all active non-manager staff
        for staff in active:
            if "manager" in staff["roles"]:
                continue
            name = staff["name"]
            if name not in self.avail_vars:
                self.avail_vars[name] = {}
                for day in DAYS:
                    saved_val = self.data.availability.get(name, {}).get(day, "off")
                    if staff.get("fixed_schedule") and saved_val == "off":
                        default_avail = staff.get("default_availability", {})
                        if default_avail and day in default_avail:
                            saved_val = default_avail[day]
                        else:
                            default_off = staff.get("default_off", [])
                            saved_val = "off" if day in default_off else "both"
                    self.avail_vars[name][day] = tk.StringVar(value=saved_val)

        # Build name lists for dropdowns (include emergency/fill-in for manual assignment)
        servers = [s for s in active if "server" in s["roles"] and "manager" not in s["roles"]]
        hosts = [s for s in active if "host" in s["roles"] and "manager" not in s["roles"]]
        servers.sort(key=lambda s: s["name"])
        hosts.sort(key=lambda s: s["name"])

        def make_name_list(staff_list):
            """Build dropdown options: plain names."""
            names = [""]
            for s in staff_list:
                names.append(s["name"])
            return names

        server_names = make_name_list(servers)
        host_names = make_name_list(hosts)

        # Cell widgets and name lists: (role, shift, day) -> cell widget / name list
        self._cell_widgets = {}
        self._slot_names = {}

        # Pre-populate slot_names from avail_vars
        def find_display_name(real_name, name_list):
            if real_name in name_list:
                return real_name
            return real_name

        for role_key, staff_list, name_list in [
            ("server", servers, server_names),
            ("host", hosts, host_names),
        ]:
            for staff in staff_list:
                name = staff["name"]
                # Dual-role staff: only show in their preferred role section
                is_dual = "server" in staff["roles"] and "host" in staff["roles"]
                if is_dual and staff.get("role_preference") and staff["role_preference"] != role_key:
                    continue
                if name not in self.avail_vars:
                    continue
                display = find_display_name(name, name_list)
                for day in DAYS:
                    val = self.avail_vars[name][day].get()
                    if val in ("morning", "both"):
                        key = (role_key, "morning", day)
                        if key not in self._slot_names:
                            self._slot_names[key] = []
                        if display not in self._slot_names[key]:
                            self._slot_names[key].append(display)
                    if val in ("night", "both"):
                        key = (role_key, "night", day)
                        if key not in self._slot_names:
                            self._slot_names[key] = []
                        if display not in self._slot_names[key]:
                            self._slot_names[key].append(display)

        def make_staff_cell(parent, names_list, full_name_list, role_key, shift, day, row, col):
            """Cell with selected names (with X to remove) and a dropdown to add."""
            cell = tk.Frame(parent, relief="groove", bd=1, bg="white", width=130)
            cell.grid(row=row, column=col, padx=2, pady=2, sticky="nsew")
            cell._names = names_list  # list of currently selected names
            cell._full_list = full_name_list
            cell._role_key = role_key
            cell._shift = shift
            cell._day = day

            def extract_name(display_val):
                if not display_val:
                    return ""
                return display_val

            def sync_to_avail():
                """Sync this cell's selections back to avail_vars."""
                # This cell's names contribute to availability
                for display in cell._names:
                    real_name = extract_name(display)
                    if real_name and real_name in self.avail_vars:
                        current = self.avail_vars[real_name][day].get()
                        if shift == "morning":
                            if current in ("night", "both"):
                                self.avail_vars[real_name][day].set("both")
                            else:
                                self.avail_vars[real_name][day].set("morning")
                        else:
                            if current in ("morning", "both"):
                                self.avail_vars[real_name][day].set("both")
                            else:
                                self.avail_vars[real_name][day].set("night")

            def refresh_cell():
                # Clear cell contents
                for w in cell.winfo_children():
                    w.destroy()

                # Show selected names with X buttons
                for i, display_name in enumerate(cell._names):
                    name_frame = tk.Frame(cell, bg="white")
                    name_frame.pack(fill=tk.X, padx=2, pady=1)

                    tk.Label(name_frame, text=display_name, font=("Helvetica", 8),
                             bg="white", anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)

                    def make_remove(idx):
                        def remove(event=None):
                            removed = cell._names.pop(idx)
                            real_name = extract_name(removed)
                            # Clear this shift from avail_vars, recalc from all cells
                            if real_name in self.avail_vars:
                                self.avail_vars[real_name][day].set("off")
                                # Re-sync all cells for this name/day
                                for key, c in self._cell_widgets.items():
                                    rk, sh, d = key
                                    if d == day:
                                        for dn in c._names:
                                            rn = extract_name(dn)
                                            if rn == real_name:
                                                cur = self.avail_vars[rn][d].get()
                                                if sh == "morning":
                                                    self.avail_vars[rn][d].set(
                                                        "both" if cur == "night" else "morning")
                                                else:
                                                    self.avail_vars[rn][d].set(
                                                        "both" if cur == "morning" else "night")
                            refresh_cell()
                        return remove

                    x_btn = tk.Label(name_frame, text="x", font=("Helvetica", 8, "bold"),
                                     fg="red", bg="white", cursor="hand2")
                    x_btn.pack(side=tk.RIGHT, padx=2)
                    x_btn.bind("<Button-1>", make_remove(i))

                # Add dropdown button
                add_frame = tk.Frame(cell, bg="white")
                add_frame.pack(fill=tk.X, padx=2, pady=(2, 1))

                add_btn = tk.Label(add_frame, text="+ Add", font=("Helvetica", 8),
                                   fg="gray", bg="#f0f0f0", relief="raised", bd=1,
                                   cursor="hand2")
                add_btn.pack(fill=tk.X)

                def open_popup(event=None):
                    popup = tk.Toplevel(parent)
                    popup.overrideredirect(True)
                    popup.attributes("-topmost", True)

                    x = add_btn.winfo_rootx()
                    y = add_btn.winfo_rooty() + add_btn.winfo_height()
                    popup.geometry(f"+{x}+{y}")

                    lb_frame = tk.Frame(popup)
                    lb_frame.pack(fill=tk.BOTH, expand=True)

                    scrollbar = tk.Scrollbar(lb_frame)
                    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

                    # Filter out already-selected names
                    selected_real = {extract_name(n) for n in cell._names}
                    available = [n for n in cell._full_list
                                 if n and extract_name(n) not in selected_real]

                    lb = tk.Listbox(lb_frame, width=18, height=min(12, max(len(available), 1)),
                                    font=("Helvetica", 9), selectmode=tk.SINGLE,
                                    yscrollcommand=scrollbar.set, activestyle="dotbox")
                    lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                    scrollbar.config(command=lb.yview)

                    for name in available:
                        lb.insert(tk.END, name)

                    def on_select(event=None):
                        sel = lb.curselection()
                        if sel:
                            chosen = lb.get(sel[0])
                            if chosen:
                                cell._names.append(chosen)
                                sync_to_avail()
                                refresh_cell()
                        popup.destroy()

                    def on_key(event):
                        ch = event.char
                        if not ch or not ch.isalpha():
                            if event.keysym == 'Return':
                                on_select()
                                return "break"
                            if event.keysym == 'Escape':
                                popup.destroy()
                                return "break"
                            return
                        ch = ch.lower()
                        for i in range(lb.size()):
                            item = lb.get(i)
                            if item and item.lower().startswith(ch):
                                lb.see(i)
                                lb.selection_clear(0, tk.END)
                                lb.selection_set(i)
                                lb.activate(i)
                                return "break"
                        return "break"

                    def on_arrow(event):
                        idx = lb.index(tk.ACTIVE)
                        if event.keysym == 'Down' and idx < lb.size() - 1:
                            idx += 1
                        elif event.keysym == 'Up' and idx > 0:
                            idx -= 1
                        lb.selection_clear(0, tk.END)
                        lb.selection_set(idx)
                        lb.activate(idx)
                        lb.see(idx)
                        return "break"

                    lb.bind("<ButtonRelease-1>", on_select)
                    lb.bind("<Key>", on_key)
                    lb.bind("<Down>", on_arrow)
                    lb.bind("<Up>", on_arrow)
                    lb.bind("<Escape>", lambda e: popup.destroy())
                    popup.update_idletasks()
                    lb.focus_force()

                    # Close popup when clicking anywhere outside it
                    def _make_close_handler(p, toplevel):
                        def handler(event):
                            try:
                                p_str = str(p)
                                w_str = str(event.widget)
                                if w_str != p_str and not w_str.startswith(p_str + "."):
                                    p.destroy()
                            except tk.TclError:
                                pass
                        def cleanup(event):
                            if event.widget == p:
                                try:
                                    toplevel.unbind("<Button-1>")
                                except tk.TclError:
                                    pass
                        return handler, cleanup
                    _toplevel = parent.winfo_toplevel()
                    _click_handler, _destroy_handler = _make_close_handler(popup, _toplevel)
                    _toplevel.bind("<Button-1>", _click_handler, add="+")
                    popup.bind("<Destroy>", _destroy_handler)

                add_btn.bind("<Button-1>", open_popup)

            refresh_cell()
            return cell

        def add_shift_section(row, title, name_list, role_key, shift, color):
            """Add a single shift-role section (e.g. Host Morning, Lunch Server)."""
            # Section header
            ttk.Label(self.avail_frame, text=title,
                      font=("Helvetica", 12, "bold"), foreground="navy").grid(
                row=row, column=0, columnspan=8, padx=5, pady=(10, 3), sticky="w")
            row += 1

            # Day column headers
            ttk.Label(self.avail_frame, text="", width=10).grid(row=row, column=0)
            for di, day in enumerate(DAYS):
                ttk.Label(self.avail_frame, text=day, font=("Helvetica", 10, "bold"),
                          width=16).grid(row=row, column=di + 1, padx=2, pady=3)
            row += 1

            # Staff cells
            ttk.Label(self.avail_frame, text=shift.capitalize(),
                      font=("Helvetica", 10, "bold"), foreground=color).grid(
                row=row, column=0, padx=5, pady=(5, 2), sticky="nw")
            for di, day in enumerate(DAYS):
                key = (role_key, shift, day)
                names = self._slot_names.get(key, [])
                cell = make_staff_cell(self.avail_frame, names, name_list,
                                       role_key, shift, day, row, di + 1)
                self._cell_widgets[key] = cell
                self._slot_names[key] = names
            row += 1

            return row

        row = 0

        # Layout matches the sign-up sheet order:
        # 1. Host Morning
        row = add_shift_section(row, "HOST — MORNING", host_names, "host", "morning", "darkorange")

        ttk.Separator(self.avail_frame, orient="horizontal").grid(
            row=row, column=0, columnspan=8, sticky="ew", pady=8)
        row += 1

        # 2. Lunch Servers
        row = add_shift_section(row, "SERVER — LUNCH", server_names, "server", "morning", "darkorange")

        ttk.Separator(self.avail_frame, orient="horizontal").grid(
            row=row, column=0, columnspan=8, sticky="ew", pady=8)
        row += 1

        # 3. Host Night
        row = add_shift_section(row, "HOST — DINNER", host_names, "host", "night", "darkblue")

        ttk.Separator(self.avail_frame, orient="horizontal").grid(
            row=row, column=0, columnspan=8, sticky="ew", pady=8)
        row += 1

        # 4. Dinner Servers
        row = add_shift_section(row, "SERVER — DINNER", server_names, "server", "night", "darkblue")

    def _set_all_available(self):
        if not messagebox.askyesno("Set All Available", "Are you sure you want to set all staff as available?"):
            return
        for staff in self.data.staff:
            if not staff["active"] or "manager" in staff["roles"]:
                continue
            name = staff["name"]
            for day in DAYS:
                self.data.availability.setdefault(name, {})[day] = "both"
        self._build_availability_grid()

    def _clear_availability(self):
        if not messagebox.askyesno("Clear All", "Are you sure you want to clear all availability?"):
            return
        for staff in self.data.staff:
            if not staff["active"] or "manager" in staff["roles"]:
                continue
            name = staff["name"]
            if staff.get("fixed_schedule"):
                default_avail = staff.get("default_availability", {})
                for day in DAYS:
                    if default_avail and day in default_avail:
                        self.data.availability.setdefault(name, {})[day] = default_avail[day]
                    else:
                        default_off = staff.get("default_off", [])
                        self.data.availability.setdefault(name, {})[day] = "off" if day in default_off else "both"
            else:
                for day in DAYS:
                    self.data.availability.setdefault(name, {})[day] = "off"
        self._build_availability_grid()

    def _save_availability(self):
        # Full sync from cell widgets to avail_vars
        if hasattr(self, '_cell_widgets'):
            def extract_name(display_val):
                if not display_val:
                    return ""
                return display_val

            # Reset all non-fixed to off
            for name, days in self.avail_vars.items():
                staff = self.data.get_staff_by_name(name)
                if staff and staff.get("fixed_schedule"):
                    continue
                for day in DAYS:
                    days[day].set("off")

            # Rebuild from cells
            for (role_key, shift, day), cell in self._cell_widgets.items():
                for display in cell._names:
                    real_name = extract_name(display)
                    if real_name and real_name in self.avail_vars:
                        current = self.avail_vars[real_name][day].get()
                        if shift == "morning":
                            if current in ("night", "both"):
                                self.avail_vars[real_name][day].set("both")
                            else:
                                self.avail_vars[real_name][day].set("morning")
                        else:
                            if current in ("morning", "both"):
                                self.avail_vars[real_name][day].set("both")
                            else:
                                self.avail_vars[real_name][day].set("night")

        avail = {}
        for name, days in self.avail_vars.items():
            avail[name] = {}
            for day, var in days.items():
                avail[name][day] = var.get()
        self.data.availability = avail
        self.data.save_availability()
        messagebox.showinfo("Saved", "Availability saved!")

    # ---- PHOTO IMPORT (OCR) ----
    def _import_from_photo(self):
        filepath = filedialog.askopenfilename(
            title="Select Sign-Up Sheet Photo",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.tiff *.webp"),
                       ("All files", "*.*")])
        if not filepath:
            return

        # Show a "processing" message
        progress_win = tk.Toplevel(self.root)
        progress_win.title("Processing...")
        progress_win.geometry("350x80")
        progress_win.transient(self.root)
        progress_win.grab_set()
        ttk.Label(progress_win, text="Reading sign-up sheet with OCR...",
                  font=("Helvetica", 11)).pack(expand=True)
        progress_win.update()

        try:
            detected, raw_ocr_texts = self._run_ocr(filepath)
        except Exception as e:
            progress_win.destroy()
            messagebox.showerror("OCR Error", f"Failed to read image:\n{e}")
            return

        progress_win.destroy()

        if not detected:
            messagebox.showwarning("No Results",
                                   "Could not detect any staff names in the image.\n"
                                   "Try a clearer photo or enter availability manually.")
            return

        self._show_ocr_confirmation(detected, raw_ocr_texts)

    # Nickname / handwriting alias map: alias -> canonical staff name
    NAME_ALIASES = {
        # Makayla
        "mak": "Makayla", "makayla": "Makayla", "mok": "Makayla",
        "makaila": "Makayla", "makayia": "Makayla",
        # Kamryn
        "kam": "Kamryn", "kamryn": "Kamryn", "kamrin": "Kamryn",
        "kamrya": "Kamryn", "kamryu": "Kamryn",
        # Dian
        "dran": "Dian", "dlan": "Dian", "diane": "Dian",
        "dion": "Dian", "dran": "Dian", "diau": "Dian",
        "dcam": "Dian", "dcan": "Dian",
        # Catie Grey
        "catie grace": "Catie Grey", "catie gray": "Catie Grey",
        "catiegrace": "Catie Grey", "catiegray": "Catie Grey",
        "catie grey": "Catie Grey", "catiegrey": "Catie Grey",
        "catie gvey": "Catie Grey", "catie guey": "Catie Grey",
        "catie": "Catie Grey", "grey": "Catie Grey",
        # Abigail
        "abigail": "Abigail", "abigai": "Abigail", "abigal": "Abigail",
        "abiguito": "Abigail", "abiguit": "Abigail", "abigui": "Abigail",
        "abiquii": "Abigail", "abiqui": "Abigail", "abiquito": "Abigail",
        "abiga": "Abigail", "abigu": "Abigail", "abigi": "Abigail",
        "abig": "Abigail", "abiqail": "Abigail", "abigoll": "Abigail",
        "abigaie": "Abigail", "abigaif": "Abigail",
        # Q
        "q": "Q", "q ": "Q",
        # Lilly
        "lilly": "Lilly", "lily": "Lilly", "lily": "Lilly",
        "lilli": "Lilly", "liily": "Lilly", "lllly": "Lilly",
        # Olivia
        "olivia": "Olivia", "olivio": "Olivia", "oivia": "Olivia",
        "oiivia": "Olivia",
        # Ashlyn
        "ashlyn": "Ashlyn", "ashlya": "Ashlyn", "ashiyn": "Ashlyn",
        "ashlyu": "Ashlyn", "ashlyh": "Ashlyn", "ashuyn": "Ashlyn",
        # Addison
        "addison": "Addison", "addisoh": "Addison", "addisoa": "Addison",
        "addisor": "Addison", "addisoo": "Addison", "addisca": "Addison",
        "addiscn": "Addison",
        # Hannah
        "hannah": "Hannah", "hannoh": "Hannah", "hannak": "Hannah",
        # Maria
        "maria": "Maria", "mario": "Maria", "marig": "Maria",
        # Leony
        "leony": "Leony", "leonv": "Leony", "leonu": "Leony",
        "leouy": "Leony", "leong": "Leony",
        # Maddie
        "maddie": "Maddie", "maddi": "Maddie", "maddre": "Maddie",
        "maddle": "Maddie", "moddre": "Maddie", "moddie": "Maddie",
        # Owen
        "owen": "Owen", "owen": "Owen", "oweu": "Owen", "owev": "Owen",
        # Sadie
        "sadie": "Sadie", "sadle": "Sadie", "sodic": "Sadie",
        "sodie": "Sadie",
        # Arabella
        "arabella": "Arabella", "arabela": "Arabella", "arobella": "Arabella",
        "arabello": "Arabella", "arahella": "Arabella", "arabela": "Arabella",
        "arabeila": "Arabella", "arabdla": "Arabella", "aradella": "Arabella",
        # Kat
        "kat": "Kat", "kot": "Kat", "kaf": "Kat",
        # Trish
        "trish": "Trish", "trisk": "Trish", "tish": "Trish",
        # Sam
        "sam": "Sam", "som": "Sam", "scm": "Sam",
        # Garret
        "garret": "Garret", "garrel": "Garret", "garref": "Garret",
        "garrett": "Garret", "garrct": "Garret", "gavret": "Garret",
        # Collin
        "collin": "Collin", "collih": "Collin", "collir": "Collin",
        "collia": "Collin", "colln": "Collin", "collia": "Collin",
        # Bryan
        "bryan": "Bryan", "bryah": "Bryan", "bryau": "Bryan",
        "bryon": "Bryan", "bryar": "Bryan", "boyar": "Bryan",
        # Jazz
        "jazz": "Jazz", "jozz": "Jazz", "jass": "Jazz", "jaz": "Jazz",
        "j2z": "Jazz", "j2zz": "Jazz", "jaet": "Jazz", "jaze": "Jazz",
        "joz": "Jazz", "jas": "Jazz",
        "ja3z": "Jazz", "j3zz": "Jazz", "j3z": "Jazz",
        "ja33": "Jazz", "j33": "Jazz", "ja3": "Jazz",
        "ja2z": "Jazz", "j2zz": "Jazz", "ja2": "Jazz", "jaz2": "Jazz",
        # Will
        "will": "Will", "wili": "Will", "wiil": "Will",
        # Andy
        "andy": "Andy", "audy": "Andy", "andv": "Andy", "ady": "Andy",
        # Winnie
        "winnie": "Winnie", "winnle": "Winnie", "wlnnie": "Winnie",
        "wianie": "Winnie", "wirmie": "Winnie",
        # Ross
        "ross": "Ross",
    }

    # Row labels that indicate morning (lunch) shifts
    MORNING_ROW_KEYWORDS = ["hostess 1", "hostess 2", "lunch server", "lunch buser",
                            "host 1", "host 2", "10:30", "10:15", "2:15", "2:30"]
    # Row labels that indicate night (dinner) shifts
    NIGHT_ROW_KEYWORDS = ["hostess 3", "dinner server", "dinner buser", "buser",
                          "host 3", "3a", "3b", "hostess 3a", "hostess 3b",
                          "4:30", "4:00", "6:00", "close", "8:30"]

    def _get_ocr_reader(self):
        """Start OCR worker subprocess (Python 3.12 venv), cache on instance."""
        if not hasattr(self, "_ocr_proc") or self._ocr_proc.poll() is not None:
            import subprocess, json
            self._ocr_proc = subprocess.Popen(
                [str(_OCR_VENV_PYTHON), str(_OCR_WORKER)],
                stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                stderr=subprocess.DEVNULL, text=True)
            # Wait for "ready" signal
            ready = json.loads(self._ocr_proc.stdout.readline())
            if ready.get("status") != "ready":
                raise RuntimeError("OCR worker failed to start")
        return self._ocr_proc

    def _paddle_readtext(self, reader_proc, img_path, **kwargs):
        """Send image to OCR worker subprocess, return (bbox, text, conf) list."""
        import json
        request = json.dumps({"cmd": "ocr", "path": str(img_path)})
        reader_proc.stdin.write(request + "\n")
        reader_proc.stdin.flush()
        response = json.loads(reader_proc.stdout.readline())
        if "error" in response:
            raise RuntimeError(f"OCR error: {response['error']}")
        return [(r[0], r[1], r[2]) for r in response["results"]]

    def _run_ocr(self, filepath):
        """Run PaddleOCR on the image using OpenCV grid detection."""
        reader = self._get_ocr_reader()

        img_cv = cv2.imread(filepath)
        if img_cv is None:
            raise ValueError(f"Could not load image: {filepath}")
        img_height, img_width = img_cv.shape[:2]

        # Build known names list
        known_names = []
        for s in self.data.staff:
            if s["active"] and "manager" not in s["roles"]:
                known_names.append(s["name"])

        # --- Grid detection with OpenCV morphological ops ---
        gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
        binary = cv2.adaptiveThreshold(
            gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY_INV, 15, 5)

        h_kernel_len = max(img_width // 12, 80)
        h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (h_kernel_len, 1))
        h_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, h_kernel, iterations=2)

        v_kernel_len = max(img_height // 12, 80)
        v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, v_kernel_len))
        v_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, v_kernel, iterations=2)

        # Find line positions from projections
        def find_line_positions(projection, threshold):
            positions = []
            in_line = False
            line_start = 0
            for i in range(len(projection)):
                if projection[i] > threshold:
                    if not in_line:
                        line_start = i
                        in_line = True
                else:
                    if in_line:
                        positions.append((line_start + i) // 2)
                        in_line = False
            if in_line:
                positions.append((line_start + len(projection)) // 2)
            return positions

        def merge_close(positions, min_gap=25):
            if not positions:
                return positions
            positions.sort()
            merged = [positions[0]]
            for p in positions[1:]:
                if p - merged[-1] < min_gap:
                    merged[-1] = (merged[-1] + p) // 2
                else:
                    merged.append(p)
            return merged

        h_positions = merge_close(find_line_positions(
            np.sum(h_lines, axis=1), img_width * 0.15))
        v_positions = merge_close(find_line_positions(
            np.sum(v_lines, axis=0), img_height * 0.15))

        if len(h_positions) < 3 or len(v_positions) < 3:
            return self._run_ocr_fallback(filepath, reader, known_names)

        # Grid line mask for removal from cell crops
        grid_mask = cv2.dilate(
            cv2.add(h_lines, v_lines), np.ones((3, 3), np.uint8), iterations=1)

        # --- Preprocessing helper ---
        def preprocess_cell(x1, y1, x2, y2):
            """Remove grid lines, CLAHE enhance, upscale 2x. Returns path to temp image."""
            cell = img_cv[y1:y2, x1:x2].copy()
            mask_crop = grid_mask[y1:y2, x1:x2]
            cell[mask_crop > 0] = [255, 255, 255]
            cell_gray = cv2.cvtColor(cell, cv2.COLOR_BGR2GRAY)
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(4, 4))
            cell_gray = clahe.apply(cell_gray)
            # Otsu binarization for cleaner black-on-white text
            _, cell_gray = cv2.threshold(cell_gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            # Upscale 3x for better text detection (especially handwriting)
            cell_up = cv2.resize(cell_gray, None, fx=3, fy=3,
                                 interpolation=cv2.INTER_CUBIC)
            import tempfile
            tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
            cv2.imwrite(tmp.name, cell_up)
            return tmp.name

        # --- Fuzzy match helpers ---
        # Single-char staff names need special handling
        single_char_names = {n.lower(): n for n in known_names if len(n) == 1}

        def fuzzy_match_name(text):
            text_clean = text.strip().lower()
            for ch in "✓✔☑▪•|[]{}()_~#@!$%^&*+=<>\"'/\\.,;:":
                text_clean = text_clean.replace(ch, "")
            # Strip trailing/leading digits (checkmarks misread as numbers)
            text_clean = text_clean.strip("0123456789 ")
            # Also try alias with embedded digits removed (e.g. "j2z" -> "jz")
            text_no_digits = "".join(c for c in text_clean if not c.isdigit())
            if len(text_clean) < 1:
                return None

            # Check for single-character names (like "Q") anywhere in text
            if single_char_names:
                # Exact single char match
                if text_clean in single_char_names:
                    return single_char_names[text_clean]
                # Single char surrounded by spaces or at word boundaries
                for ch, name in single_char_names.items():
                    words = text_clean.split()
                    if ch in words:
                        return name

            # Exact alias (also try with digits removed)
            if text_clean in self.NAME_ALIASES:
                return self.NAME_ALIASES[text_clean]
            if text_no_digits and text_no_digits != text_clean and text_no_digits in self.NAME_ALIASES:
                return self.NAME_ALIASES[text_no_digits]
            # Alias substring
            for alias, canonical in self.NAME_ALIASES.items():
                if len(alias) >= 3 and alias in text_clean:
                    return canonical

            # Fuzzy match
            best_name = None
            best_score = 0.0
            for name in known_names:
                if len(name) == 1:
                    continue  # Already handled above
                nl = name.lower()
                score = SequenceMatcher(None, text_clean, nl).ratio()
                if nl in text_clean or text_clean in nl:
                    score = max(score, 0.88)
                if len(text_clean) >= 3 and len(nl) >= 3 and text_clean[:3] == nl[:3]:
                    score = max(score, 0.75)
                if score > best_score:
                    best_score = score
                    best_name = name

            threshold = 0.7 if len(text_clean) <= 3 else 0.45
            return best_name if best_score >= threshold else None

        day_keywords = {
            "mon": "MON", "monday": "MON",
            "tue": "TUE", "tues": "TUE", "tuesday": "TUE",
            "wed": "WED", "wednesday": "WED",
            "thu": "THU", "thur": "THU", "thurs": "THU", "thursday": "THU",
            "fri": "FRI", "friday": "FRI",
            "sat": "SAT", "saturday": "SAT",
            "sun": "SUN", "sunday": "SUN",
        }

        def fuzzy_match_day(text):
            text_clean = text.strip().lower().replace(".", "").replace("/", " ")
            for keyword, day in day_keywords.items():
                if keyword in text_clean:
                    return day
            return None

        def detect_shift(label_text):
            lt = label_text.lower()
            is_morning = any(kw in lt for kw in self.MORNING_ROW_KEYWORDS)
            is_night = any(kw in lt for kw in self.NIGHT_ROW_KEYWORDS)
            if is_morning and not is_night:
                return "morning"
            if is_night and not is_morning:
                return "night"
            return None

        def detect_row_role(label_text):
            """Detect whether a row is for hosts or servers."""
            lt = label_text.lower()
            if any(kw in lt for kw in ["hostess", "host"]):
                return "host"
            if any(kw in lt for kw in ["server", "buser"]):
                return "server"
            return None

        # Build role lookup for staff
        staff_roles = {}
        for s in self.data.staff:
            if s["active"]:
                staff_roles[s["name"]] = set(s.get("roles", []))

        # --- Step 1: Detect day columns from header ---
        # Run PaddleOCR on header area to find day names
        day_for_col = {}

        # Find the header row: first row with printed day names
        # Try rows until we find one with >= 5 days detected
        for header_ri in range(min(3, len(h_positions) - 1)):
            if len(day_for_col) >= 5:
                break
            hy1 = h_positions[header_ri]
            hy2 = h_positions[header_ri + 1] if header_ri + 1 < len(h_positions) else img_height
            for ci in range(len(v_positions) - 1):
                if ci in day_for_col:
                    continue
                cx1, cx2 = v_positions[ci] + 3, v_positions[ci + 1] - 3
                if cx2 - cx1 < 20:
                    continue
                tmp_path = preprocess_cell(cx1, hy1 + 3, cx2, hy2 - 3)
                results = self._paddle_readtext(reader, tmp_path)
                os.unlink(tmp_path)
                for _, text, _ in results:
                    day = fuzzy_match_day(text)
                    if day and day not in day_for_col.values():
                        day_for_col[ci] = day
                        break

        # --- Step 2: Detect row labels and shifts ---
        row_shifts = {}  # row_index -> "morning" / "night" / None
        row_roles = {}   # row_index -> "host" / "server" / None
        for ri in range(len(h_positions) - 1):
            ry1, ry2 = h_positions[ri], h_positions[ri + 1]
            if ry2 - ry1 < 15:
                continue
            # Label is in the first grid column (between first two vertical lines)
            lx1 = v_positions[0] + 3
            lx2 = v_positions[1] - 3 if len(v_positions) > 1 else img_width // 8
            tmp_path = preprocess_cell(lx1, ry1 + 3, lx2, ry2 - 3)
            results = self._paddle_readtext(reader, tmp_path)
            os.unlink(tmp_path)
            label_text = " ".join(t for _, t, _ in results)
            shift = detect_shift(label_text)
            if shift:
                row_shifts[ri] = shift
            role = detect_row_role(label_text)
            if role:
                row_roles[ri] = role

        # --- Step 3: OCR each data cell ---
        detected_names = {}  # name -> {day -> set of shifts}
        raw_ocr_texts = {}   # name -> set of raw OCR strings that matched

        for ri in range(len(h_positions) - 1):
            ry1, ry2 = h_positions[ri], h_positions[ri + 1]
            if ry2 - ry1 < 30:
                continue
            row_shift = row_shifts.get(ri)
            row_role = row_roles.get(ri)  # "host", "server", or None
            # Skip header rows (no shift detected and near top)
            if row_shift is None and ri < 2:
                continue

            for ci, day in day_for_col.items():
                if ci >= len(v_positions) - 1:
                    continue
                cx1, cx2 = v_positions[ci] + 3, v_positions[ci + 1] - 3
                if cx2 - cx1 < 20:
                    continue

                tmp_path = preprocess_cell(cx1, ry1 + 3, cx2, ry2 - 3)
                results = self._paddle_readtext(reader, tmp_path)
                os.unlink(tmp_path)

                for _, text, conf in results:
                    if not text or len(text.strip()) < 1:
                        continue
                    if conf < 0.5:
                        continue
                    # Split on common separators: &, "and", newlines, slashes
                    parts = [text]
                    for sep in ["\n", " & ", " and ", "&", "/"]:
                        new_parts = []
                        for p in parts:
                            new_parts.extend([s.strip() for s in p.split(sep) if s.strip()])
                        parts = new_parts

                    # Also try splitting long text that may contain multiple names
                    # e.g., "Leony Dian" -> ["Leony", "Dian"]
                    expanded = []
                    for part in parts:
                        expanded.append(part)
                        # If no direct match, try splitting on spaces
                        if not fuzzy_match_name(part) and " " in part:
                            words = part.split()
                            # Try each word individually
                            expanded.extend(words)
                            # Try consecutive word pairs
                            for i in range(len(words) - 1):
                                expanded.append(words[i] + " " + words[i + 1])
                    parts = expanded

                    for part in parts:
                        name = fuzzy_match_name(part)
                        if name:
                            # Skip if this row's role doesn't match the staff member's roles
                            if row_role and name in staff_roles:
                                if row_role not in staff_roles[name]:
                                    continue
                            if name not in detected_names:
                                detected_names[name] = {}
                            if name not in raw_ocr_texts:
                                raw_ocr_texts[name] = set()
                            role_tag = f"[{row_role}]" if row_role else ""
                            raw_ocr_texts[name].add(f"{part.strip()} {role_tag}".strip())
                            shift = row_shift or "both"
                            if day not in detected_names[name]:
                                detected_names[name][day] = set()
                            detected_names[name][day].add(shift)

        # Build result: merge shifts per day
        result = {}
        for name in known_names:
            if name in detected_names and detected_names[name]:
                avail = {}
                for day in DAYS:
                    if day in detected_names[name]:
                        shifts = detected_names[name][day]
                        if "morning" in shifts and "night" in shifts:
                            avail[day] = "both"
                        elif "both" in shifts:
                            avail[day] = "both"
                        elif "morning" in shifts:
                            avail[day] = "morning"
                        elif "night" in shifts:
                            avail[day] = "night"
                        else:
                            avail[day] = "both"
                    else:
                        avail[day] = "off"
                result[name] = avail

        return result, raw_ocr_texts

    def _run_ocr_fallback(self, filepath, reader, known_names):
        """Fallback OCR when grid detection fails — run on full image."""
        results = self._paddle_readtext(reader, filepath)
        detected_names = {}
        raw_ocr_texts = {}
        for _, text, conf in results:
            if conf < 0.4:
                continue
            text_clean = text.strip().lower()
            for ch in "✓✔☑▪•|[]{}()_~#@!$%^&*+=<>\"'/\\.,;:":
                text_clean = text_clean.replace(ch, "")
            text_clean = text_clean.strip("0123456789 ")
            if len(text_clean) < 2:
                continue
            if text_clean in self.NAME_ALIASES:
                name = self.NAME_ALIASES[text_clean]
            else:
                best_name = None
                best_score = 0.0
                for n in known_names:
                    score = SequenceMatcher(None, text_clean, n.lower()).ratio()
                    if n.lower() in text_clean or text_clean in n.lower():
                        score = max(score, 0.85)
                    if score > best_score:
                        best_score = score
                        best_name = n
                name = best_name if best_score >= 0.6 else None
            if name:
                if name not in detected_names:
                    detected_names[name] = set(DAYS)
                if name not in raw_ocr_texts:
                    raw_ocr_texts[name] = set()
                raw_ocr_texts[name].add(text.strip())

        result = {}
        for name in known_names:
            if name in detected_names:
                result[name] = {day: "both" for day in DAYS}
        return result, raw_ocr_texts

    def _show_ocr_confirmation(self, detected, raw_ocr_texts=None):
        """Show a dialog for the manager to review and confirm OCR results — cell-based layout."""
        if raw_ocr_texts is None:
            raw_ocr_texts = {}
        dialog = tk.Toplevel(self.root)
        dialog.title("Confirm Sign-Up Sheet")
        dialog.geometry("1100x650")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Review detected availability from photo",
                  font=("Helvetica", 12, "bold")).pack(pady=(10, 3))
        ttk.Label(dialog, text="Add/remove names in each cell, then click Apply.",
                  font=("Helvetica", 10)).pack(pady=(0, 8))

        # Scrollable frame
        canvas = tk.Canvas(dialog)
        v_scroll = ttk.Scrollbar(dialog, orient=tk.VERTICAL, command=canvas.yview)
        h_scroll = ttk.Scrollbar(dialog, orient=tk.HORIZONTAL, command=canvas.xview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)

        canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

        # Build staff groups
        active = [s for s in self.data.staff if s["active"]
                  and "manager" not in s["roles"]]
        servers = [s for s in active if "server" in s["roles"]]
        hosts = [s for s in active if "host" in s["roles"]]
        servers.sort(key=lambda s: s["name"])
        hosts.sort(key=lambda s: s["name"])

        def make_name_list(staff_list):
            names = []
            for s in staff_list:
                names.append(s["name"])
            return names

        server_names = make_name_list(servers)
        host_names = make_name_list(hosts)

        def extract_name(display_val):
            if not display_val:
                return ""
            parts = display_val.rsplit(" ", 1)
            if len(parts) == 2 and parts[1] in ("server", "host", "dual"):
                return parts[0]
            return display_val

        def find_display(real_name, name_list):
            for n in name_list:
                if extract_name(n) == real_name:
                    return n
            return real_name

        # Build cell data from detected results
        # ocr_cells: (role, shift, day) -> [display_name, ...]
        ocr_cells = {}
        ocr_cell_widgets = {}

        for role_key, staff_list, name_list in [("server", servers, server_names),
                                                 ("host", hosts, host_names)]:
            for staff in staff_list:
                name = staff["name"]
                display = find_display(name, name_list)
                # Dual-role staff: only show in their preferred role section
                is_dual = "server" in staff["roles"] and "host" in staff["roles"]
                if is_dual and staff.get("role_preference") and staff["role_preference"] != role_key:
                    continue
                # Fixed schedule staff use their set schedule
                if staff.get("fixed_schedule"):
                    default_avail = staff.get("default_availability", {})
                    default_off = staff.get("default_off", [])
                    for day in DAYS:
                        if default_avail and day in default_avail:
                            val = default_avail[day]
                        else:
                            val = "off" if day in default_off else "both"
                        if val in ("morning", "both"):
                            key = (role_key, "morning", day)
                            ocr_cells.setdefault(key, [])
                            if display not in ocr_cells[key]:
                                ocr_cells[key].append(display)
                        if val in ("night", "both"):
                            key = (role_key, "night", day)
                            ocr_cells.setdefault(key, [])
                            if display not in ocr_cells[key]:
                                ocr_cells[key].append(display)
                elif name in detected:
                    for day in DAYS:
                        val = detected[name].get(day, "off")
                        if val in ("morning", "both"):
                            key = (role_key, "morning", day)
                            ocr_cells.setdefault(key, [])
                            if display not in ocr_cells[key]:
                                ocr_cells[key].append(display)
                        if val in ("night", "both"):
                            key = (role_key, "night", day)
                            ocr_cells.setdefault(key, [])
                            if display not in ocr_cells[key]:
                                ocr_cells[key].append(display)

        def make_ocr_cell(parent, names_list, full_name_list, row, col):
            """Cell with detected names (X to remove) and + Add dropdown."""
            cell = tk.Frame(parent, relief="groove", bd=1, bg="white", width=130)
            cell.grid(row=row, column=col, padx=2, pady=2, sticky="nsew")
            cell._names = names_list
            cell._full_list = full_name_list

            def refresh():
                for w in cell.winfo_children():
                    w.destroy()
                for i, display_name in enumerate(cell._names):
                    nf = tk.Frame(cell, bg="white")
                    nf.pack(fill=tk.X, padx=2, pady=1)
                    real = extract_name(display_name)
                    # Green if detected, blue if fixed schedule
                    color = "green" if real in detected else "blue"
                    tk.Label(nf, text=display_name, font=("Helvetica", 8),
                             bg="white", fg=color, anchor="w").pack(
                        side=tk.LEFT, fill=tk.X, expand=True)
                    def make_remove(idx):
                        def remove(event=None):
                            cell._names.pop(idx)
                            refresh()
                        return remove
                    x_btn = tk.Label(nf, text="x", font=("Helvetica", 8, "bold"),
                                     fg="red", bg="white", cursor="hand2")
                    x_btn.pack(side=tk.RIGHT, padx=2)
                    x_btn.bind("<Button-1>", make_remove(i))

                # Raw OCR info
                detected_in_cell = [extract_name(n) for n in cell._names if extract_name(n) in raw_ocr_texts]
                if detected_in_cell:
                    raw_texts = []
                    for n in detected_in_cell:
                        raw_texts.extend(raw_ocr_texts[n])
                    raw_str = ", ".join(sorted(set(raw_texts)))
                    if raw_str:
                        tk.Label(cell, text=f"OCR: {raw_str}", font=("Helvetica", 7),
                                 fg="gray", bg="white", wraplength=120, anchor="w").pack(
                            fill=tk.X, padx=2)

                # Add button
                add_frame = tk.Frame(cell, bg="white")
                add_frame.pack(fill=tk.X, padx=2, pady=(2, 1))
                add_btn = tk.Label(add_frame, text="+ Add", font=("Helvetica", 8),
                                   fg="gray", bg="#f0f0f0", relief="raised", bd=1,
                                   cursor="hand2")
                add_btn.pack(fill=tk.X)

                def open_popup(event=None):
                    popup = tk.Toplevel(parent)
                    popup.overrideredirect(True)
                    popup.attributes("-topmost", True)
                    x = add_btn.winfo_rootx()
                    y = add_btn.winfo_rooty() + add_btn.winfo_height()
                    popup.geometry(f"+{x}+{y}")

                    selected_real = {extract_name(n) for n in cell._names}
                    available = [n for n in cell._full_list if n and extract_name(n) not in selected_real]

                    lb_frame = tk.Frame(popup)
                    lb_frame.pack(fill=tk.BOTH, expand=True)
                    scrollbar = tk.Scrollbar(lb_frame)
                    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                    lb = tk.Listbox(lb_frame, width=18, height=min(12, max(len(available), 1)),
                                    font=("Helvetica", 9), selectmode=tk.SINGLE,
                                    yscrollcommand=scrollbar.set, activestyle="dotbox")
                    lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                    scrollbar.config(command=lb.yview)
                    for name in available:
                        lb.insert(tk.END, name)

                    def on_select(event=None):
                        sel = lb.curselection()
                        if sel:
                            chosen = lb.get(sel[0])
                            if chosen:
                                cell._names.append(chosen)
                                refresh()
                        popup.destroy()

                    def on_key(event):
                        ch = event.char
                        if not ch or not ch.isalpha():
                            if event.keysym == 'Return':
                                on_select()
                                return "break"
                            if event.keysym == 'Escape':
                                popup.destroy()
                                return "break"
                            return
                        ch = ch.lower()
                        for i in range(lb.size()):
                            item = lb.get(i)
                            if item and item.lower().startswith(ch):
                                lb.see(i)
                                lb.selection_clear(0, tk.END)
                                lb.selection_set(i)
                                lb.activate(i)
                                return "break"
                        return "break"

                    def on_arrow(event):
                        idx = lb.index(tk.ACTIVE)
                        if event.keysym == 'Down' and idx < lb.size() - 1:
                            idx += 1
                        elif event.keysym == 'Up' and idx > 0:
                            idx -= 1
                        lb.selection_clear(0, tk.END)
                        lb.selection_set(idx)
                        lb.activate(idx)
                        lb.see(idx)
                        return "break"

                    lb.bind("<ButtonRelease-1>", on_select)
                    lb.bind("<Key>", on_key)
                    lb.bind("<Down>", on_arrow)
                    lb.bind("<Up>", on_arrow)
                    lb.bind("<Escape>", lambda e: popup.destroy())
                    popup.update_idletasks()
                    lb.focus_force()

                    # Close popup when clicking anywhere outside it
                    def _make_close_handler(p, toplevel):
                        def handler(event):
                            try:
                                p_str = str(p)
                                w_str = str(event.widget)
                                if w_str != p_str and not w_str.startswith(p_str + "."):
                                    p.destroy()
                            except tk.TclError:
                                pass
                        def cleanup(event):
                            if event.widget == p:
                                try:
                                    toplevel.unbind("<Button-1>")
                                except tk.TclError:
                                    pass
                        return handler, cleanup
                    _toplevel = parent.winfo_toplevel()
                    _click_handler, _destroy_handler = _make_close_handler(popup, _toplevel)
                    _toplevel.bind("<Button-1>", _click_handler, add="+")
                    popup.bind("<Destroy>", _destroy_handler)

                add_btn.bind("<Button-1>", open_popup)

            refresh()
            return cell

        def add_shift_section(row, title, name_list, role_key, shift, color):
            """Add a single shift-role section to the OCR confirmation."""
            ttk.Label(inner, text=title,
                      font=("Helvetica", 12, "bold"), foreground="navy").grid(
                row=row, column=0, columnspan=8, padx=5, pady=(10, 3), sticky="w")
            row += 1

            ttk.Label(inner, text="", width=10).grid(row=row, column=0)
            for di, day in enumerate(DAYS):
                ttk.Label(inner, text=day, font=("Helvetica", 10, "bold"),
                          width=16).grid(row=row, column=di + 1, padx=2, pady=3)
            row += 1

            ttk.Label(inner, text=shift.capitalize(),
                      font=("Helvetica", 10, "bold"), foreground=color).grid(
                row=row, column=0, padx=5, pady=(5, 2), sticky="nw")
            for di, day in enumerate(DAYS):
                key = (role_key, shift, day)
                names = ocr_cells.get(key, [])
                cell = make_ocr_cell(inner, names, name_list, row, di + 1)
                ocr_cell_widgets[key] = cell
            row += 1

            return row

        row = 0

        # Layout matches sign-up sheet order
        row = add_shift_section(row, "HOST — MORNING", host_names, "host", "morning", "darkorange")

        ttk.Separator(inner, orient="horizontal").grid(
            row=row, column=0, columnspan=8, sticky="ew", pady=8)
        row += 1

        row = add_shift_section(row, "SERVER — LUNCH", server_names, "server", "morning", "darkorange")

        ttk.Separator(inner, orient="horizontal").grid(
            row=row, column=0, columnspan=8, sticky="ew", pady=8)
        row += 1

        row = add_shift_section(row, "HOST — DINNER", host_names, "host", "night", "darkblue")

        ttk.Separator(inner, orient="horizontal").grid(
            row=row, column=0, columnspan=8, sticky="ew", pady=8)
        row += 1

        row = add_shift_section(row, "SERVER — DINNER", server_names, "server", "night", "darkblue")

        # Buttons
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)

        def apply_ocr():
            # Build availability from cell contents
            for name in [s["name"] for s in active]:
                staff = self.data.get_staff_by_name(name)
                if staff and staff.get("fixed_schedule"):
                    continue
                self.data.availability.setdefault(name, {})
                for day in DAYS:
                    self.data.availability[name][day] = "off"

            for (role_key, shift, day), cell in ocr_cell_widgets.items():
                for display in cell._names:
                    real_name = extract_name(display)
                    if not real_name:
                        continue
                    self.data.availability.setdefault(real_name, {})
                    current = self.data.availability[real_name].get(day, "off")
                    if shift == "morning":
                        if current in ("night", "both"):
                            self.data.availability[real_name][day] = "both"
                        else:
                            self.data.availability[real_name][day] = "morning"
                    else:
                        if current in ("morning", "both"):
                            self.data.availability[real_name][day] = "both"
                        else:
                            self.data.availability[real_name][day] = "night"

            dialog.destroy()
            self._build_availability_grid()
            messagebox.showinfo("Applied",
                                "Availability updated from photo.\n"
                                "Click 'Save Availability' to save.")

        def reset_to_detected():
            # Rebuild ocr_cells from detected data and refresh all cells
            for (role_key, shift, day), cell in ocr_cell_widgets.items():
                key = (role_key, shift, day)
                original = ocr_cells.get(key, [])
                cell._names.clear()
                cell._names.extend(original)
                # Refresh cell display
                for w in cell.winfo_children():
                    w.destroy()
                # Trigger refresh by re-calling make_ocr_cell logic
            # Easier: just rebuild the whole dialog content
            dialog.destroy()
            self._show_ocr_confirmation(detected, raw_ocr_texts)

        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="Apply", command=apply_ocr).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="Reset to Detected", command=reset_to_detected).pack(side=tk.LEFT, padx=5)

    # ---- TAB 4: SCHEDULE GENERATION ----
    def _build_schedule_tab(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="  Schedule  ")

        # Toolbar
        toolbar = ttk.Frame(frame)
        toolbar.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(toolbar, text="Week Starting (Mon):").pack(side=tk.LEFT, padx=5)
        self.week_start_var = tk.StringVar(value=self.data.config.get("week_start", ""))
        ttk.Entry(toolbar, textvariable=self.week_start_var, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Label(toolbar, text="(YYYY-MM-DD)").pack(side=tk.LEFT, padx=2)

        ttk.Button(toolbar, text="Generate Schedule",
                  command=self._generate_schedule).pack(side=tk.LEFT, padx=15)
        ttk.Button(toolbar, text="Select Midshifts",
                  command=self._select_midshifts).pack(side=tk.LEFT, padx=5)

        # Warning area
        self.warning_var = tk.StringVar(value="")
        self.warning_label = ttk.Label(frame, textvariable=self.warning_var,
                                       foreground="red", wraplength=1100)
        self.warning_label.pack(fill=tk.X, padx=10)

        # Schedule display - scrollable canvas with a table
        canvas_frame = ttk.Frame(frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.sched_canvas = tk.Canvas(canvas_frame)
        v_scroll = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.sched_canvas.yview)
        h_scroll = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.sched_canvas.xview)
        self.sched_inner = ttk.Frame(self.sched_canvas)
        self.sched_inner.bind("<Configure>",
                              lambda e: self.sched_canvas.configure(scrollregion=self.sched_canvas.bbox("all")))
        self.sched_canvas.create_window((0, 0), window=self.sched_inner, anchor="nw")
        self.sched_canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        self.sched_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)

    def _generate_schedule(self):
        # Save availability first
        self._save_availability_silent()
        self._save_config_silent()

        self.current_schedule, self.current_warnings = self.generator.generate()

        if self.current_warnings:
            self.warning_var.set("Warnings: " + " | ".join(self.current_warnings))
        else:
            self.warning_var.set("Schedule generated successfully!")

        self._display_schedule()

    def _save_availability_silent(self):
        if hasattr(self, 'avail_vars'):
            avail = {}
            for name, days in self.avail_vars.items():
                avail[name] = {}
                for day, var in days.items():
                    avail[name][day] = var.get()
            self.data.availability = avail
            self.data.save_availability()

    def _save_config_silent(self):
        if hasattr(self, 'config_vars'):
            keys = ["lunch_servers", "lunch_hosts", "hibachi_lunch",
                    "dinner_servers", "dinner_hosts", "hibachi_dinner"]
            for day in DAYS:
                for key in keys:
                    if key in self.config_vars[day]:
                        self.data.config["staffing"][day][key] = self.config_vars[day][key].get()
            self.data.config["week_start"] = self.week_start_var.get()
            self.data.save_config()

    def _display_schedule(self):
        for widget in self.sched_inner.winfo_children():
            widget.destroy()

        if not self.current_schedule:
            return

        # Header row
        ttk.Label(self.sched_inner, text="", width=18, font=("Helvetica", 9, "bold")).grid(
            row=0, column=0, padx=2, pady=2)
        for di, day in enumerate(DAYS):
            lbl = ttk.Label(self.sched_inner, text=day, width=18,
                           font=("Helvetica", 10, "bold"), anchor="center")
            lbl.grid(row=0, column=di + 1, padx=2, pady=2)

        row = 1
        # Section: LUNCH SERVERS
        ttk.Label(self.sched_inner, text="LUNCH SERVERS",
                 font=("Helvetica", 9, "bold"), foreground="blue").grid(
            row=row, column=0, padx=5, pady=3, sticky="w")
        ttk.Label(self.sched_inner, text=LUNCH_SERVER_TIME,
                 font=("Helvetica", 8)).grid(row=row + 1, column=0, padx=5, sticky="w")
        row += 1

        max_ls = max(len(self.current_schedule[d]["lunch_servers"]) + len(self.current_schedule[d]["hibachi_lunch_servers"]) for d in DAYS)
        for si in range(max(max_ls, 1)):
            for di, day in enumerate(DAYS):
                hibachi_lunch = self.current_schedule[day]["hibachi_lunch_servers"]
                servers = self.current_schedule[day]["lunch_servers"]
                letters = self.current_schedule[day].get("lunch_server_letters", {})
                hib_set = set(hibachi_lunch)
                # Sort by letter, hibachi after regular within same letter
                combined = servers + hibachi_lunch
                combined.sort(key=lambda n: (letters.get(n, "Z"), 1 if n in hib_set else 0))
                mids = self.current_schedule[day].get("mid_servers", [])
                if si < len(combined):
                    name = combined[si]
                    letter = letters.get(name, chr(65 + si))
                    mid = "*" if name in mids else ""
                    is_hib = name in hibachi_lunch
                    text = f"{name} {letter}{mid}"
                    if is_hib:
                        bg = "#FFE4B5"
                    else:
                        bg = "#E8F5E9" if mid else "white"
                    lbl = tk.Label(self.sched_inner, text=text, width=18,
                                  font=("Helvetica", 9), anchor="center",
                                  bg=bg, relief="groove", bd=1, cursor="hand2")
                    lbl.grid(row=row + si, column=di + 1, padx=1, pady=1)
                    lbl.bind("<Button-1>", lambda e, d=day, k="lunch_servers", i=si: self._edit_cell(d, k, i))
        row += max(max_ls, 1) + 1

        # Lunch Manager
        ttk.Label(self.sched_inner, text="LUNCH MANAGER",
                 font=("Helvetica", 9, "bold"), foreground="darkgreen").grid(
            row=row, column=0, padx=5, pady=3, sticky="w")
        for di, day in enumerate(DAYS):
            mgr = self.current_schedule[day]["lunch_manager"]
            lbl = tk.Label(self.sched_inner, text=mgr, width=18,
                          font=("Helvetica", 9, "bold"), anchor="center",
                          relief="groove", bd=1)
            lbl.grid(row=row, column=di + 1, padx=1, pady=1)
        row += 1

        # Lunch Hosts
        ttk.Label(self.sched_inner, text="LUNCH HOSTS",
                 font=("Helvetica", 9, "bold"), foreground="teal").grid(
            row=row, column=0, padx=5, pady=3, sticky="w")
        row += 1
        max_lh = max(len(self.current_schedule[d]["lunch_hosts"]) for d in DAYS)
        for hi in range(max(max_lh, 1)):
            for di, day in enumerate(DAYS):
                hosts = self.current_schedule[day]["lunch_hosts"]
                if hi < len(hosts):
                    lbl = tk.Label(self.sched_inner, text=hosts[hi], width=18,
                                  font=("Helvetica", 9), anchor="center",
                                  relief="groove", bd=1, cursor="hand2")
                    lbl.grid(row=row + hi, column=di + 1, padx=1, pady=1)
                    lbl.bind("<Button-1>", lambda e, d=day, k="lunch_hosts", i=hi: self._edit_cell(d, k, i))
        row += max(max_lh, 1) + 1

        # Separator
        ttk.Separator(self.sched_inner, orient="horizontal").grid(
            row=row, column=0, columnspan=8, sticky="ew", pady=5)
        row += 1

        # DINNER SERVERS
        ttk.Label(self.sched_inner, text="DINNER SERVERS",
                 font=("Helvetica", 9, "bold"), foreground="darkred").grid(
            row=row, column=0, padx=5, pady=3, sticky="w")
        ttk.Label(self.sched_inner, text=DINNER_SERVER_TIME,
                 font=("Helvetica", 8)).grid(row=row + 1, column=0, padx=5, sticky="w")
        row += 1

        max_ds = max(
            len(self.current_schedule[d]["dinner_servers"]) +
            len(self.current_schedule[d]["hibachi_servers"])
            for d in DAYS)
        for si in range(max(max_ds, 1)):
            for di, day in enumerate(DAYS):
                hibachi = self.current_schedule[day]["hibachi_servers"]
                servers = self.current_schedule[day]["dinner_servers"]
                letters = self.current_schedule[day].get("dinner_server_letters", {})
                hib_set = set(hibachi)
                # Sort by letter, hibachi after regular within same letter
                combined = servers + hibachi
                combined.sort(key=lambda n: (letters.get(n, "Z"), 1 if n in hib_set else 0))
                mids = self.current_schedule[day].get("mid_servers", [])
                if si < len(combined):
                    name = combined[si]
                    letter = letters.get(name, chr(65 + si))
                    mid = "*" if name in mids else ""
                    is_hib = name in hibachi
                    text = f"{name} {letter}{mid}"
                    if is_hib:
                        bg = "#FFE4B5"
                    else:
                        bg = "#E8F5E9" if mid else "white"
                    lbl = tk.Label(self.sched_inner, text=text, width=18,
                                  font=("Helvetica", 9), anchor="center",
                                  bg=bg, relief="groove", bd=1, cursor="hand2")
                    lbl.grid(row=row + si, column=di + 1, padx=1, pady=1)
                    lbl.bind("<Button-1>", lambda e, d=day, k="dinner_servers", i=si: self._edit_cell(d, k, i))
        row += max(max_ds, 1) + 1

        # Dinner Hosts
        ttk.Label(self.sched_inner, text="DINNER HOSTS",
                 font=("Helvetica", 9, "bold"), foreground="teal").grid(
            row=row, column=0, padx=5, pady=3, sticky="w")
        row += 1
        max_dh = max(len(self.current_schedule[d]["dinner_hosts"]) for d in DAYS)
        for hi in range(max(max_dh, 1)):
            for di, day in enumerate(DAYS):
                hosts = self.current_schedule[day]["dinner_hosts"]
                if hi < len(hosts):
                    lbl = tk.Label(self.sched_inner, text=hosts[hi], width=18,
                                  font=("Helvetica", 9), anchor="center",
                                  relief="groove", bd=1, cursor="hand2")
                    lbl.grid(row=row + hi, column=di + 1, padx=1, pady=1)
                    lbl.bind("<Button-1>", lambda e, d=day, k="dinner_hosts", i=hi: self._edit_cell(d, k, i))
        row += max(max_dh, 1) + 1

        # Dinner Manager
        ttk.Label(self.sched_inner, text="DINNER MANAGER",
                 font=("Helvetica", 9, "bold"), foreground="darkgreen").grid(
            row=row, column=0, padx=5, pady=3, sticky="w")
        for di, day in enumerate(DAYS):
            mgr = self.current_schedule[day]["dinner_manager"]
            lbl = tk.Label(self.sched_inner, text=mgr, width=18,
                          font=("Helvetica", 9, "bold"), anchor="center",
                          relief="groove", bd=1)
            lbl.grid(row=row, column=di + 1, padx=1, pady=1)
        row += 2

        # Shift count summary
        ttk.Label(self.sched_inner, text="SHIFT TOTALS",
                 font=("Helvetica", 9, "bold")).grid(row=row, column=0, padx=5, pady=5, sticky="w")
        row += 1
        counts = self.generator.shift_counts
        sorted_counts = sorted(counts.items(), key=lambda x: x[1], reverse=True)
        col_offset = 0
        for i, (name, count) in enumerate(sorted_counts):
            r = row + (i % 10)
            c = (i // 10) * 2
            ttk.Label(self.sched_inner, text=f"{name}: {count} shifts",
                     font=("Helvetica", 8)).grid(row=r, column=c, padx=5, sticky="w")

    def _get_available_for_slot(self, day, key):
        """Build list of available staff for a schedule slot, with role labels for dual-role."""
        shift_type = "morning" if "lunch" in key else "night"
        slot_role = "server" if "server" in key else "host"

        avail = self.data.availability
        options = []

        for s in self.data.staff:
            if not s["active"] or "manager" in s["roles"]:
                continue
            name = s["name"]

            # Check availability for this day/shift
            staff_avail = avail.get(name, {}).get(day, "off")
            if staff_avail == "off":
                continue
            if shift_type == "morning" and staff_avail == "night":
                continue
            if shift_type == "night" and staff_avail == "morning":
                continue

            # Check role match and label dual-role staff
            is_dual = "server" in s["roles"] and "host" in s["roles"]
            has_role = slot_role in s["roles"]

            if has_role and is_dual:
                options.append((name, name))
            elif has_role:
                options.append((name, name))

        options.sort(key=lambda x: x[1].lower())
        return options

    def _edit_cell(self, day, key, index):
        """Let manager swap a staff member in a cell via searchable dropdown."""
        if not self.current_schedule:
            return

        if key == "dinner_servers":
            hibachi = self.current_schedule[day]["hibachi_servers"]
            servers = self.current_schedule[day]["dinner_servers"]
            combined = hibachi + servers
            if index >= len(combined):
                return
            current_name = combined[index]
        else:
            lst = self.current_schedule[day][key]
            if index >= len(lst):
                return
            current_name = lst[index]

        # Get available staff for this slot
        available = self._get_available_for_slot(day, key)
        display_names = ["(remove)"] + [label for _, label in available]
        name_map = {"(remove)": "(remove)"}
        for real_name, label in available:
            name_map[label] = real_name

        dlg = tk.Toplevel(self.root)
        dlg.title(f"Edit {day} - {key.replace('_', ' ').title()}")
        dlg.geometry("350x150")
        dlg.transient(self.root)
        dlg.grab_set()

        ttk.Label(dlg, text=f"Current: {current_name}",
                 font=("Helvetica", 11, "bold")).pack(pady=(15, 5))
        ttk.Label(dlg, text="Type to search, then select:").pack()

        combo_var = tk.StringVar()
        combo = ttk.Combobox(dlg, textvariable=combo_var, values=display_names,
                             width=35, font=("Helvetica", 10))
        combo.pack(padx=15, pady=10)
        combo.focus_set()

        # Filter dropdown as user types
        def on_keyrelease(event):
            if event.keysym in ('Return', 'Tab', 'Escape', 'Up', 'Down', 'Left', 'Right'):
                return
            typed = combo_var.get().lower()
            if not typed:
                combo["values"] = display_names
                return
            filtered = [n for n in display_names if n.lower().startswith(typed)]
            if not filtered:
                filtered = [n for n in display_names if typed in n.lower()]
            combo["values"] = filtered if filtered else display_names

        combo.bind("<KeyRelease>", on_keyrelease)
        combo.bind("<FocusIn>", lambda e: combo.configure(values=display_names))

        def apply(event=None):
            selection = combo_var.get()
            if not selection:
                dlg.destroy()
                return
            new_name = name_map.get(selection, selection)

            if key == "dinner_servers":
                hibachi = self.current_schedule[day]["hibachi_servers"]
                servers = self.current_schedule[day]["dinner_servers"]
                if index < len(hibachi):
                    if new_name == "(remove)":
                        hibachi.pop(index)
                    else:
                        hibachi[index] = new_name
                else:
                    adj_idx = index - len(hibachi)
                    if new_name == "(remove)":
                        servers.pop(adj_idx)
                    else:
                        servers[adj_idx] = new_name
            else:
                lst = self.current_schedule[day][key]
                if new_name == "(remove)":
                    lst.pop(index)
                else:
                    lst[index] = new_name

            self._display_schedule()
            dlg.destroy()

        combo.bind("<<ComboboxSelected>>", apply)
        ttk.Button(dlg, text="Apply", command=apply).pack(pady=5)

    def _select_midshifts(self):
        """Open dialog for manager to pick midshift servers"""
        if not self.current_schedule:
            messagebox.showinfo("Info", "Generate a schedule first.")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("Select Midshift Servers")
        dlg.geometry("700x500")
        dlg.transient(self.root)
        dlg.grab_set()

        ttk.Label(dlg, text="Check servers for midshift (* mark)",
                 font=("Helvetica", 11, "bold")).pack(pady=10)

        canvas = tk.Canvas(dlg)
        scrollbar = ttk.Scrollbar(dlg, orient=tk.VERTICAL, command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Headers
        ttk.Label(inner, text="Server", width=12, font=("Helvetica", 9, "bold")).grid(
            row=0, column=0, padx=5, pady=3)
        for di, day in enumerate(DAYS):
            ttk.Label(inner, text=day, width=8, font=("Helvetica", 9, "bold")).grid(
                row=0, column=di + 1, padx=3, pady=3)

        mid_vars = {}
        # Get all servers who appear in any shift
        all_names = set()
        for day in DAYS:
            for name in self.current_schedule[day]["lunch_servers"]:
                all_names.add(name)
            for name in self.current_schedule[day]["hibachi_lunch_servers"]:
                all_names.add(name)
            for name in self.current_schedule[day]["dinner_servers"]:
                all_names.add(name)
            for name in self.current_schedule[day]["hibachi_servers"]:
                all_names.add(name)

        # Dian and Leony have fixed midshifts
        existing_mids = {}
        for day in DAYS:
            existing_mids[day] = self.current_schedule[day].get("mid_servers", [])

        for ri, name in enumerate(sorted(all_names)):
            ttk.Label(inner, text=name, width=12).grid(row=ri + 1, column=0, padx=5, pady=2, sticky="w")
            mid_vars[name] = {}
            for di, day in enumerate(DAYS):
                # Check if this person works this day
                works = (name in self.current_schedule[day]["lunch_servers"] or
                         name in self.current_schedule[day]["hibachi_lunch_servers"] or
                         name in self.current_schedule[day]["dinner_servers"] or
                         name in self.current_schedule[day]["hibachi_servers"])
                is_mid = name in existing_mids.get(day, [])
                var = tk.BooleanVar(value=is_mid)
                mid_vars[name][day] = var
                cb = ttk.Checkbutton(inner, variable=var)
                if not works:
                    cb.configure(state="disabled")
                # Dian/Leony midshifts are fixed
                staff = self.data.get_staff_by_name(name)
                if staff and staff.get("fixed_schedule"):
                    cb.configure(state="disabled")
                cb.grid(row=ri + 1, column=di + 1, padx=3, pady=2)

        def apply_mids():
            for day in DAYS:
                mids = []
                for name, days in mid_vars.items():
                    if day in days and days[day].get():
                        mids.append(name)
                self.current_schedule[day]["mid_servers"] = mids
            self._display_schedule()
            dlg.destroy()

        ttk.Button(dlg, text="Apply Midshifts", command=apply_mids).pack(pady=10)

    # ---- TAB 5: EXPORT ----
    def _build_export_tab(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="  Export  ")

        ttk.Label(frame, text="Export Schedule",
                 font=("Helvetica", 14, "bold")).pack(pady=20)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=20)

        ttk.Button(btn_frame, text="Export to Excel (.xlsx)",
                  command=self._export_excel).pack(pady=10, ipadx=20, ipady=10)
        ttk.Button(btn_frame, text="Export to PDF",
                  command=self._export_pdf).pack(pady=10, ipadx=20, ipady=10)

        # Status
        self.export_status = tk.StringVar(value="Generate a schedule first, then export.")
        ttk.Label(frame, textvariable=self.export_status,
                 font=("Helvetica", 10)).pack(pady=20)

    def _get_week_start_date(self):
        date_str = self.week_start_var.get().strip()
        if date_str:
            try:
                return datetime.strptime(date_str, "%Y-%m-%d")
            except ValueError:
                pass
        return None

    def _export_excel(self):
        if not self.current_schedule:
            messagebox.showinfo("Info", "Generate a schedule first.")
            return

        week_date = self._get_week_start_date()
        default_name = "Schedule"
        if week_date:
            end_date = week_date + timedelta(days=6)
            default_name = f"{week_date.strftime('%Y %b %d')} - {end_date.strftime('%d')}"

        filepath = filedialog.asksaveasfilename(
            title="Save Excel Schedule",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialdir=self.data.config.get("export_path", "") or str(APP_DIR),
            initialfile=default_name,
        )
        if not filepath:
            return

        # Save the export path for next time
        self.data.config["export_path"] = str(Path(filepath).parent)
        self.data.save_config()

        try:
            exporter = ExcelExporter(self.current_schedule, self.data.config, week_date)
            exporter.export(filepath)
            self.export_status.set(f"Excel saved: {filepath}")
            messagebox.showinfo("Success", f"Schedule exported to:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed:\n{e}")

    def _export_pdf(self):
        if not self.current_schedule:
            messagebox.showinfo("Info", "Generate a schedule first.")
            return

        if FPDF is None:
            messagebox.showerror("Error", "fpdf2 not installed. Run setup.sh to install.")
            return

        week_date = self._get_week_start_date()
        default_name = "Schedule"
        if week_date:
            end_date = week_date + timedelta(days=6)
            default_name = f"{week_date.strftime('%Y %b %d')} - {end_date.strftime('%d')}"

        filepath = filedialog.asksaveasfilename(
            title="Save PDF Schedule",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialdir=self.data.config.get("export_path", "") or str(APP_DIR),
            initialfile=default_name,
        )
        if not filepath:
            return

        self.data.config["export_path"] = str(Path(filepath).parent)
        self.data.save_config()

        try:
            exporter = PDFExporter(self.current_schedule, self.data.config, week_date)
            exporter.export(filepath)
            self.export_status.set(f"PDF saved: {filepath}")
            messagebox.showinfo("Success", f"Schedule exported to:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed:\n{e}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    root = tk.Tk()
    app = ToyoSchedulerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
