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

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
DAYS = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
SHIFTS = ["morning", "mid", "night"]

MORNING_SIDEWORK = ["A", "B", "C", "D"]
MORNING_EXTRA_ORDER = ["A", "C", "D"]  # doubles in this order
NIGHT_SIDEWORK = ["A", "B", "C", "D", "E", "F", "G", "H"]
NIGHT_EXTRA_ORDER = ["F", "H"]  # doubles in this order

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
    {"name": "Diane", "roles": ["server"], "seniority": 10, "fixed_schedule": True,
     "active": True, "flags": [], "role_preference": "server",
     "default_off": ["MON", "TUE"]},
    {"name": "Leony", "roles": ["server"], "seniority": 10, "fixed_schedule": True,
     "active": True, "flags": [], "role_preference": "server",
     "default_off": ["WED", "THU"]},
    {"name": "Will", "roles": ["server"], "seniority": 5, "fixed_schedule": False,
     "active": True, "flags": ["always_hibachi"], "role_preference": "server"},
    {"name": "Andy", "roles": ["server"], "seniority": 6, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "server",
     "preferred_off": ["TUE", "WED", "THU"]},
    {"name": "Winnie", "roles": ["server"], "seniority": 3, "fixed_schedule": False,
     "active": True, "flags": ["emergency_only"], "role_preference": "server"},
    {"name": "Cindy", "roles": ["server"], "seniority": 3, "fixed_schedule": False,
     "active": False, "flags": ["fill_in"], "role_preference": "server"},
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
    {"name": "Catie Grace", "roles": ["host"], "seniority": 4, "fixed_schedule": False,
     "active": True, "flags": [], "role_preference": "host"},
    {"name": "Olivia", "roles": ["host"], "seniority": 8, "fixed_schedule": False,
     "active": True, "flags": ["seniority_priority"], "role_preference": "host"},
    {"name": "Lily", "roles": ["host"], "seniority": 4, "fixed_schedule": False,
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
                "hibachi_servers": [],
            }
            for day in DAYS
        }
        self.shift_counts = {}  # name -> total shifts
        self.hibachi_counts = {}  # name -> hibachi shifts this week
        self.warnings = []

    def generate(self):
        self.shift_counts = {}
        self.hibachi_counts = {}
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

        # Step 1: Place managers
        for day in DAYS:
            self.schedule[day]["lunch_manager"] = config["managers"].get("lunch", "Aaron")
            self.schedule[day]["dinner_manager"] = config["managers"].get("dinner", "Aaron")

        # Step 2: Place fixed schedule staff (Diane & Leony)
        self._place_fixed_staff(avail, config)

        # Step 3: Place Will (always hibachi)
        self._place_will(avail, config)

        # Step 4: Place hosts with seniority (Olivia first)
        self._place_hosts(avail, config)

        # Step 5: Place remaining servers
        self._place_servers(avail, config)

        # Step 6: Assign sidework letters
        self._assign_sidework_letters()

        # Step 7: Balance hibachi (ensure ~2 per server per week)
        self._balance_hibachi(config)

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

    def _is_assigned(self, name, day):
        s = self.schedule[day]
        for key in ["lunch_servers", "lunch_hosts", "mid_servers",
                     "dinner_servers", "dinner_hosts", "hibachi_servers"]:
            if name in s[key]:
                return True
        return False

    def _add_shift(self, name):
        self.shift_counts[name] = self.shift_counts.get(name, 0) + 1

    def _get_shift_count(self, name):
        return self.shift_counts.get(name, 0)

    def _place_fixed_staff(self, avail, config):
        for name in ["Diane", "Leony"]:
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
                          and not self._is_assigned(h["name"], day)]
            # Sort: fewer shifts first (balance), then seniority for ties
            candidates.sort(key=lambda h: (self._get_shift_count(h["name"]), -h["seniority"]))
            for h in candidates[:needed - len(self.schedule[day]["lunch_hosts"])]:
                self.schedule[day]["lunch_hosts"].append(h["name"])
                self._add_shift(h["name"])

            # Dinner hosts
            needed = staffing["dinner_hosts"]
            candidates = [h for h in hosts
                          if self._can_work(h["name"], day, "night")
                          and not self._is_assigned(h["name"], day)
                          and not ("no_closing" in (self.data.get_staff_by_name(h["name"]) or {}).get("flags", []))]
            candidates.sort(key=lambda h: (self._get_shift_count(h["name"]), -h["seniority"]))
            for h in candidates[:needed - len(self.schedule[day]["dinner_hosts"])]:
                self.schedule[day]["dinner_hosts"].append(h["name"])
                self._add_shift(h["name"])

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
                            and not self._is_assigned(s["name"], day)]
                    dual.sort(key=lambda s: self._get_shift_count(s["name"]))
                    for s in dual[:needed - current]:
                        self.schedule[day][shift_key].append(s["name"])
                        self._add_shift(s["name"])

    def _place_servers(self, avail, config):
        all_servers = [s for s in self.data.get_active_staff("server")
                       if not s["fixed_schedule"]
                       and "always_hibachi" not in s["flags"]
                       and "fill_in" not in s["flags"]
                       and "emergency_only" not in s["flags"]]

        fill_ins = [s for s in self.data.get_active_staff("server")
                    if "fill_in" in s["flags"]]

        for day in DAYS:
            # Sort fill-ins: deprioritize those who prefer this day off
            def fill_in_sort_key(s):
                preferred_off = s.get("preferred_off", [])
                penalty = 10 if day in preferred_off else 0
                return (penalty, self._get_shift_count(s["name"]))

            staffing = config["staffing"][day]

            # Lunch servers
            needed = staffing["lunch_servers"]
            current = len(self.schedule[day]["lunch_servers"])
            if current < needed:
                candidates = [s for s in all_servers
                              if self._can_work(s["name"], day, "morning")
                              and not self._is_assigned(s["name"], day)]
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
                        and not self._is_assigned(s["name"], day)]
                dual.sort(key=lambda s: self._get_shift_count(s["name"]))
                for s in dual[:needed - current]:
                    self.schedule[day]["lunch_servers"].append(s["name"])
                    self._add_shift(s["name"])

            # Still short? Try fill-ins
            current = len(self.schedule[day]["lunch_servers"])
            if current < needed:
                fi = [s for s in fill_ins
                      if self._can_work(s["name"], day, "morning")
                      and not self._is_assigned(s["name"], day)]
                fi.sort(key=fill_in_sort_key)
                for s in fi[:needed - current]:
                    self.schedule[day]["lunch_servers"].append(s["name"])
                    self._add_shift(s["name"])

            # Dinner servers
            needed = staffing["dinner_servers"]
            current = len(self.schedule[day]["dinner_servers"])
            if current < needed:
                candidates = [s for s in all_servers
                              if self._can_work(s["name"], day, "night")
                              and not self._is_assigned(s["name"], day)]
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
                        and not self._is_assigned(s["name"], day)]
                dual.sort(key=lambda s: self._get_shift_count(s["name"]))
                for s in dual[:needed - current]:
                    self.schedule[day]["dinner_servers"].append(s["name"])
                    self._add_shift(s["name"])

            # Dinner - try fill-ins
            current = len(self.schedule[day]["dinner_servers"])
            if current < needed:
                fi = [s for s in fill_ins
                      if self._can_work(s["name"], day, "night")
                      and not self._is_assigned(s["name"], day)]
                fi.sort(key=fill_in_sort_key)
                for s in fi[:needed - current]:
                    self.schedule[day]["dinner_servers"].append(s["name"])
                    self._add_shift(s["name"])

            # Midshift
            needed_mid = staffing["mid_servers"]
            # Mid servers are selected by manager later, but we can suggest
            # For now leave empty - manager picks in the GUI

            # Check shortages
            for key, label in [("lunch_servers", "Lunch Servers"),
                               ("dinner_servers", "Dinner Servers"),
                               ("lunch_hosts", "Lunch Hosts"),
                               ("dinner_hosts", "Dinner Hosts")]:
                needed_key = key
                if needed_key == "lunch_hosts":
                    n = staffing["lunch_hosts"]
                elif needed_key == "dinner_hosts":
                    n = staffing["dinner_hosts"]
                elif needed_key == "lunch_servers":
                    n = staffing["lunch_servers"]
                else:
                    n = staffing["dinner_servers"]
                actual = len(self.schedule[day][key])
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
        #   Fri/Sat: Andy -> Diane/Leony -> Abigail -> Bryan -> random
        #   Weekdays: Diane -> Leony -> Abigail -> Bryan -> random
        #   Sun: Diane -> Leony -> Abigail -> Bryan -> random
        MORNING_A_PRIORITY_FRISATAM = ["Andy", "Diane", "Leony", "Abigail", "Bryan"]
        MORNING_A_PRIORITY_DEFAULT = ["Diane", "Leony", "Abigail", "Bryan"]

        # Night letter-A priority:
        #   Andy -> Diane -> Leony -> Sadie -> Bryan -> random
        NIGHT_A_PRIORITY = ["Andy", "Diane", "Leony", "Sadie", "Bryan"]

        for day in DAYS:
            # --- Morning servers get A-D ---
            servers = list(self.schedule[day]["lunch_servers"])
            letters = list(MORNING_SIDEWORK)
            if len(servers) > len(letters):
                for extra_letter in MORNING_EXTRA_ORDER:
                    letters.append(extra_letter)
                    if len(letters) >= len(servers):
                        break

            self.schedule[day]["lunch_server_letters"] = {}

            # Determine morning A holder
            if day in ("FRI", "SAT"):
                morning_a = self._find_a_holder(servers, MORNING_A_PRIORITY_FRISATAM)
            else:
                morning_a = self._find_a_holder(servers, MORNING_A_PRIORITY_DEFAULT)

            if morning_a:
                ordered = [morning_a] + [n for n in servers if n != morning_a]
            else:
                ordered = servers

            for i, name in enumerate(ordered):
                if i < len(letters):
                    self.schedule[day]["lunch_server_letters"][name] = letters[i]
                else:
                    self.schedule[day]["lunch_server_letters"][name] = letters[-1]

            # --- Night servers get A-H ---
            all_night = self.schedule[day]["dinner_servers"] + self.schedule[day]["hibachi_servers"]
            letters = list(NIGHT_SIDEWORK)
            if len(all_night) > len(letters):
                for extra_letter in NIGHT_EXTRA_ORDER:
                    letters.append(extra_letter)
                    if len(letters) >= len(all_night):
                        break

            self.schedule[day]["dinner_server_letters"] = {}

            night_a = self._find_a_holder(all_night, NIGHT_A_PRIORITY)

            if night_a:
                ordered_night = [night_a] + [n for n in all_night if n != night_a]
            else:
                ordered_night = all_night

            for i, name in enumerate(ordered_night):
                if i < len(letters):
                    self.schedule[day]["dinner_server_letters"][name] = letters[i]
                else:
                    self.schedule[day]["dinner_server_letters"][name] = letters[-1]

    def _balance_hibachi(self, config):
        # Each server should get ~2 hibachi shifts per week
        # Will is excluded (always hibachi)
        all_servers = [s["name"] for s in self.data.get_active_staff("server")
                       if not s["fixed_schedule"]
                       and "always_hibachi" not in s["flags"]
                       and "fill_in" not in s["flags"]
                       and "emergency_only" not in s["flags"]
                       and s["active"]]

        # Track who has hibachi assignments
        for name in all_servers:
            self.hibachi_counts[name] = 0

        # For days that need hibachi servers, pick from dinner servers
        # First pass: give hibachi to those with 0 shifts
        for day in DAYS:
            staffing = config["staffing"][day]
            needed = staffing.get("hibachi_dinner", 0)
            if needed <= 0:
                continue
            dinner = self.schedule[day]["dinner_servers"][:]
            dinner_eligible = [n for n in dinner if n in all_servers
                               and self.hibachi_counts.get(n, 0) < 2]
            dinner_eligible.sort(key=lambda n: self.hibachi_counts.get(n, 0))

            picked = dinner_eligible[:needed]
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
            servers = self.schedule[day]["lunch_servers"]
            letters = self.schedule[day].get("lunch_server_letters", {})
            mids = self.schedule[day].get("mid_servers", [])
            for si, name in enumerate(servers):
                row = 3 + si
                letter = letters.get(name, chr(65 + si))
                mid_mark = "*" if name in mids else ""
                ws[f"{col}{row}"] = f"{name} {letter}{mid_mark}"
                ws[f"{col}{row}"].font = name_font
                ws[f"{col}{row}"].alignment = Alignment(horizontal="center")

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

            # Place hibachi first (in angle brackets)
            row = 21
            for name in hibachi:
                letter = letters.get(name, "W")
                mid_mark = "*" if name in mids else ""
                ws[f"{col}{row}"] = f"<{name} {letter}>{mid_mark}"
                ws[f"{col}{row}"].font = name_font
                ws[f"{col}{row}"].alignment = Alignment(horizontal="center")
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
        ws["B41"] = "Hibachi Servers"
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

        # Find max lunch servers across days
        max_lunch = max(len(self.schedule[d]["lunch_servers"]) for d in DAYS) if DAYS else 0
        for si in range(max(max_lunch, 1)):
            row = [""]
            if si == 0:
                row[0] = LUNCH_SERVER_TIME
            for day in DAYS:
                servers = self.schedule[day]["lunch_servers"]
                letters = self.schedule[day].get("lunch_server_letters", {})
                mids = self.schedule[day].get("mid_servers", [])
                if si < len(servers):
                    name = servers[si]
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

        # Lunch hosts
        max_lh = max(len(self.schedule[d]["lunch_hosts"]) for d in DAYS) if DAYS else 0
        for hi in range(max(max_lh, 1)):
            row = [f"Hostess {hi+1}" if hi < 3 else ""]
            for day in DAYS:
                hosts = self.schedule[day]["lunch_hosts"]
                row.append(hosts[hi] if hi < len(hosts) else "")
            y = draw_row(y, row, size=6)

        y += 2  # spacer

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
                combined = hibachi + servers
                letters = self.schedule[day].get("dinner_server_letters", {})
                mids = self.schedule[day].get("mid_servers", [])
                if si < len(combined):
                    name = combined[si]
                    letter = letters.get(name, chr(65 + si))
                    mid = "*" if name in mids else ""
                    prefix = "<" if name in hibachi else ""
                    suffix = ">" if name in hibachi else ""
                    row.append(f"{prefix}{name} {letter}{suffix}{mid}")
                else:
                    row.append("")
            y = draw_row(y, row, size=6)

        # Dinner hosts
        y += 1
        max_dh = max(len(self.schedule[d]["dinner_hosts"]) for d in DAYS) if DAYS else 0
        for hi in range(max(max_dh, 1)):
            row = [f"Hostess {hi+1}" if hi < 3 else ""]
            for day in DAYS:
                hosts = self.schedule[day]["dinner_hosts"]
                row.append(hosts[hi] if hi < len(hosts) else "")
            y = draw_row(y, row, size=6)

        # Dinner manager
        row_mgr = ["Dinner Manager"]
        for day in DAYS:
            row_mgr.append(self.schedule[day]["dinner_manager"])
        y = draw_row(y, row_mgr, bold=True, size=6)

        # Legend
        y += 3
        pdf.set_font("Helvetica", "I", 7)
        pdf.set_xy(x_start, y)
        pdf.cell(0, 4, "* = Midshift server(s)  |  < > = Hibachi side  |  Letters = Sidework assignment")

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

        # Manager settings
        mgr_frame = ttk.LabelFrame(frame, text="Managers")
        mgr_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(mgr_frame, text="Lunch Manager:").grid(row=0, column=0, padx=10, pady=5)
        self.lunch_mgr_var = tk.StringVar(value=self.data.config["managers"]["lunch"])
        ttk.Entry(mgr_frame, textvariable=self.lunch_mgr_var, width=15).grid(row=0, column=1, padx=5)

        ttk.Label(mgr_frame, text="Dinner Manager:").grid(row=0, column=2, padx=10, pady=5)
        self.dinner_mgr_var = tk.StringVar(value=self.data.config["managers"]["dinner"])
        ttk.Entry(mgr_frame, textvariable=self.dinner_mgr_var, width=15).grid(row=0, column=3, padx=5)

        # Staffing grid
        grid_frame = ttk.LabelFrame(frame, text="Staffing Numbers Per Day")
        grid_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        labels = ["", "Lunch\nServers", "Lunch\nHosts", "Mid\nServers",
                  "Dinner\nServers", "Dinner\nHosts", "Hibachi\nDinner"]
        for i, label in enumerate(labels):
            ttk.Label(grid_frame, text=label, font=("Helvetica", 9, "bold"),
                     justify="center").grid(row=0, column=i, padx=8, pady=5)

        self.config_vars = {}
        keys = ["lunch_servers", "lunch_hosts", "mid_servers",
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
        self.data.config["managers"]["lunch"] = self.lunch_mgr_var.get()
        self.data.config["managers"]["dinner"] = self.dinner_mgr_var.get()
        keys = ["lunch_servers", "lunch_hosts", "mid_servers",
                "dinner_servers", "dinner_hosts", "hibachi_dinner"]
        for day in DAYS:
            for key in keys:
                self.data.config["staffing"][day][key] = self.config_vars[day][key].get()
        self.data.save_config()
        messagebox.showinfo("Saved", "Configuration saved!")

    def _copy_weekday(self):
        """Copy Monday values to Tue-Thu"""
        keys = ["lunch_servers", "lunch_hosts", "mid_servers",
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

        # Headers
        ttk.Label(self.avail_frame, text="Staff", font=("Helvetica", 10, "bold"),
                 width=15).grid(row=0, column=0, padx=5, pady=3)
        ttk.Label(self.avail_frame, text="Role", font=("Helvetica", 10, "bold"),
                 width=10).grid(row=0, column=1, padx=5, pady=3)
        for di, day in enumerate(DAYS):
            ttk.Label(self.avail_frame, text=day, font=("Helvetica", 10, "bold"),
                     width=10).grid(row=0, column=di + 2, padx=3, pady=3)

        self.avail_vars = {}
        options = ["off", "morning", "night", "both"]

        active = [s for s in self.data.staff if s["active"]
                  and "emergency_only" not in s.get("flags", [])]
        # Sort: fixed schedule first, then by name
        active.sort(key=lambda s: (0 if s.get("fixed_schedule") else 1, s["name"]))

        for ri, staff in enumerate(active):
            name = staff["name"]
            ttk.Label(self.avail_frame, text=name, width=15).grid(
                row=ri + 1, column=0, padx=5, pady=2, sticky="w")
            role_str = "/".join(staff["roles"])
            ttk.Label(self.avail_frame, text=role_str, width=10).grid(
                row=ri + 1, column=1, padx=5, pady=2)

            self.avail_vars[name] = {}
            for di, day in enumerate(DAYS):
                saved_val = self.data.availability.get(name, {}).get(day, "off")
                # Default fixed-schedule staff to "both", except their default off days
                if staff.get("fixed_schedule") and saved_val == "off":
                    default_off = staff.get("default_off", [])
                    if day in default_off:
                        saved_val = "off"
                    else:
                        saved_val = "both"
                var = tk.StringVar(value=saved_val)
                self.avail_vars[name][day] = var
                combo = ttk.Combobox(self.avail_frame, textvariable=var,
                                    values=options, state="readonly", width=8)
                combo.grid(row=ri + 1, column=di + 2, padx=3, pady=2)

    def _set_all_available(self):
        for name, days in self.avail_vars.items():
            for day, var in days.items():
                var.set("both")

    def _clear_availability(self):
        for name, days in self.avail_vars.items():
            staff = self.data.get_staff_by_name(name)
            for day, var in days.items():
                if staff and staff.get("fixed_schedule"):
                    var.set("both")
                else:
                    var.set("off")

    def _save_availability(self):
        avail = {}
        for name, days in self.avail_vars.items():
            avail[name] = {}
            for day, var in days.items():
                avail[name][day] = var.get()
        self.data.availability = avail
        self.data.save_availability()
        messagebox.showinfo("Saved", "Availability saved!")

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
            self.data.config["managers"]["lunch"] = self.lunch_mgr_var.get()
            self.data.config["managers"]["dinner"] = self.dinner_mgr_var.get()
            keys = ["lunch_servers", "lunch_hosts", "mid_servers",
                    "dinner_servers", "dinner_hosts", "hibachi_dinner"]
            for day in DAYS:
                for key in keys:
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

        max_ls = max(len(self.current_schedule[d]["lunch_servers"]) for d in DAYS)
        for si in range(max(max_ls, 1)):
            for di, day in enumerate(DAYS):
                servers = self.current_schedule[day]["lunch_servers"]
                letters = self.current_schedule[day].get("lunch_server_letters", {})
                mids = self.current_schedule[day].get("mid_servers", [])
                if si < len(servers):
                    name = servers[si]
                    letter = letters.get(name, chr(65 + si))
                    mid = "*" if name in mids else ""
                    text = f"{name} {letter}{mid}"
                    fg = "purple" if mid else "black"
                    lbl = tk.Label(self.sched_inner, text=text, width=18,
                                  fg=fg, font=("Helvetica", 9), anchor="center",
                                  relief="groove", bd=1, cursor="hand2")
                    lbl.grid(row=row + si, column=di + 1, padx=1, pady=1)
                    lbl.bind("<Button-1>", lambda e, d=day, k="lunch_servers", i=si: self._edit_cell(d, k, i))
            row_label = f"Server {si + 1}" if si > 0 else ""
            if si > 0:
                ttk.Label(self.sched_inner, text=row_label).grid(row=row + si, column=0, padx=5, sticky="w")
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
                combined = hibachi + servers
                letters = self.current_schedule[day].get("dinner_server_letters", {})
                mids = self.current_schedule[day].get("mid_servers", [])
                if si < len(combined):
                    name = combined[si]
                    letter = letters.get(name, chr(65 + si))
                    mid = "*" if name in mids else ""
                    is_hib = name in hibachi
                    if is_hib:
                        text = f"<{name} {letter}>{mid}"
                        bg = "#FFE4B5"
                    else:
                        text = f"{name} {letter}{mid}"
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

    def _edit_cell(self, day, key, index):
        """Let manager swap a staff member in a cell"""
        if not self.current_schedule:
            return

        if key == "dinner_servers":
            # Combined list
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

        # Show popup to select replacement
        dlg = tk.Toplevel(self.root)
        dlg.title(f"Edit {day} - {key}")
        dlg.geometry("300x400")
        dlg.transient(self.root)
        dlg.grab_set()

        ttk.Label(dlg, text=f"Current: {current_name}",
                 font=("Helvetica", 11, "bold")).pack(pady=10)
        ttk.Label(dlg, text="Select replacement:").pack()

        listbox = tk.Listbox(dlg, width=30, height=15)
        listbox.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

        # Get eligible staff
        role = "server" if "server" in key else "host"
        eligible = [s["name"] for s in self.data.get_active_staff(role)]
        eligible.append("(remove)")
        for name in sorted(eligible):
            listbox.insert(tk.END, name)

        def apply():
            sel = listbox.curselection()
            if not sel:
                dlg.destroy()
                return
            new_name = listbox.get(sel[0])

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

        ttk.Button(dlg, text="Apply", command=apply).pack(pady=10)

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
            for name in self.current_schedule[day]["dinner_servers"]:
                all_names.add(name)
            for name in self.current_schedule[day]["hibachi_servers"]:
                all_names.add(name)

        # Diane and Leony have fixed midshifts
        existing_mids = {}
        for day in DAYS:
            existing_mids[day] = self.current_schedule[day].get("mid_servers", [])

        for ri, name in enumerate(sorted(all_names)):
            ttk.Label(inner, text=name, width=12).grid(row=ri + 1, column=0, padx=5, pady=2, sticky="w")
            mid_vars[name] = {}
            for di, day in enumerate(DAYS):
                # Check if this person works this day
                works = (name in self.current_schedule[day]["lunch_servers"] or
                         name in self.current_schedule[day]["dinner_servers"] or
                         name in self.current_schedule[day]["hibachi_servers"])
                is_mid = name in existing_mids.get(day, [])
                var = tk.BooleanVar(value=is_mid)
                mid_vars[name][day] = var
                cb = ttk.Checkbutton(inner, variable=var)
                if not works:
                    cb.configure(state="disabled")
                # Diane/Leony midshifts are fixed
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
