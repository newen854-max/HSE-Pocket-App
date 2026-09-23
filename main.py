__version__ = "2.0.0"

import os
import csv
import json
import shutil
import sqlite3
import zipfile
from pathlib import Path
from datetime import datetime

from kivy.app import App
from kivy.metrics import dp
from kivy.core.window import Window
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.spinner import Spinner
from kivy.uix.popup import Popup
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.clock import Clock

try:
    from plyer import filechooser
except Exception:
    filechooser = None

try:
    import xlsxwriter
except Exception:
    xlsxwriter = None


# ============================================================
# APPLICATION CONSTANTS
# ============================================================

APP_NAME = "HSE Management System"
APP_VERSION = __version__


# ============================================================
# APPLICATION STORAGE
# ============================================================

BASE = Path(App.user_data_dir)

DB_FILE = BASE / "hse.db"
ATTACH_DIR = BASE / "attachments"
EXPORT_DIR = BASE / "exports"
BACKUP_DIR = BASE / "backups"

for folder in (BASE, ATTACH_DIR, EXPORT_DIR, BACKUP_DIR):
    folder.mkdir(parents=True, exist_ok=True)


# ============================================================
# MASTER LISTS
# ============================================================

OBS_TYPES = [
    "Unsafe Act",
    "Unsafe Condition",
    "Good Observation",
    "Good Practice",
    "Positive Observation",
    "Environmental Observation",
    "Near Miss",
    "Other",
]

CATEGORIES = [
    "PPE",
    "Work at Height",
    "Excavation",
    "Electrical Safety",
    "Lifting Operations",
    "Scaffolding",
    "Confined Space",
    "Fire Safety",
    "Housekeeping",
    "Slip Trip Fall",
    "Chemical Safety",
    "Environmental",
    "Traffic Management",
    "Plant and Machinery",
    "Manual Handling",
    "Emergency Preparedness",
    "Heat Stress",
    "Welfare",
    "Dropped Objects",
    "Tools and Equipment",
    "Permit to Work",
    "Barricading",
    "Material Storage",
    "Access and Egress",
    "Gas Testing",
    "Hot Work",
    "Noise",
    "Dust",
    "Waste Management",
    "Environmental Spill",
    "Other",
]

PRIORITIES = [
    "Low",
    "Medium",
    "High",
    "Critical",
]

STATUSES = [
    "Open",
    "In Progress",
    "Pending Verification",
    "Closed",
    "Cancelled",
]

INCIDENT_TYPES = [
    "Fatality",
    "Lost Time Injury",
    "Medical Treatment Case",
    "Restricted Work Case",
    "First Aid Case",
    "Near Miss",
    "Property Damage",
    "Environmental Incident",
    "Fire Incident",
    "Vehicle Incident",
    "Equipment Damage",
    "Chemical Spill",
    "Other",
]

INVESTIGATION_METHODS = [
    "Root Cause Analysis",
    "5 Why Analysis",
    "Fishbone / Ishikawa",
    "ICAM",
    "Barrier Analysis",
    "Bow-Tie Analysis",
    "Fault Tree Analysis",
    "Causal Tree",
    "Other",
]

CAPA_SOURCES = [
    "HSE Inspection",
    "Incident",
    "Near Miss",
    "Audit",
    "Client Observation",
    "Regulatory Inspection",
    "Environmental Incident",
    "Management Review",
    "Employee Observation",
    "Other",
]


# ============================================================
# GENERAL HELPERS
# ============================================================

def now():
    return datetime.now()


def today():
    return now().strftime("%Y-%m-%d")


def safe(value):
    return "" if value is None else str(value)


# ============================================================
# DATABASE
# ============================================================

class DB:

    def __init__(self):
        self.conn = sqlite3.connect(str(DB_FILE))
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys=ON")
        self.init()

    def init(self):

        self.conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS settings(
                key TEXT PRIMARY KEY,
                value TEXT
            );

            CREATE TABLE IF NOT EXISTS observations(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                number TEXT UNIQUE,
                obs_date TEXT,
                obs_time TEXT,
                project TEXT,
                company TEXT,
                location TEXT,
                area TEXT,
                responsible TEXT,
                designation TEXT,
                employee_id TEXT,
                observer TEXT,
                observer_designation TEXT,
                observer_id TEXT,
                observer_company TEXT,
                obs_type TEXT,
                category TEXT,
                subcategory TEXT,
                observation TEXT,
                immediate_action TEXT,
                corrective_action TEXT,
                preventive_action TEXT,
                priority TEXT,
                target_date TEXT,
                status TEXT DEFAULT 'Open',
                closeout_comments TEXT,
                created_at TEXT
            );

            CREATE TABLE IF NOT EXISTS observation_attachments(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                observation_id INTEGER,
                file_path TEXT,
                attachment_type TEXT,
                original_name TEXT,
                FOREIGN KEY(observation_id)
                    REFERENCES observations(id)
                    ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS incidents(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                number TEXT UNIQUE,
                incident_date TEXT,
                incident_time TEXT,
                location TEXT,
                project TEXT,
                company TEXT,
                department TEXT,
                activity TEXT,
                incident_type TEXT,
                person_involved TEXT,
                employee_id TEXT,
                designation TEXT,
                supervisor TEXT,
                witnesses TEXT,
                description TEXT,
                immediate_action TEXT,
                consequences TEXT,
                potential_consequences TEXT,
                equipment TEXT,
                investigation_method TEXT,
                investigation_details TEXT,
                status TEXT DEFAULT 'Open',
                created_at TEXT
            );

            /*
             * IMPORTANT:
             *
             * ICAM timeline is stored in a separate table.
             *
             * Each event receives a permanent sequence_no:
             *
             * 1
             * 2
             * 3
             * 4
             *
             * Adding a new event therefore does NOT replace
             * or overwrite an earlier event.
             */
            CREATE TABLE IF NOT EXISTS incident_timeline(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                incident_id INTEGER NOT NULL,
                sequence_no INTEGER NOT NULL,
                event_date TEXT,
                event_time TEXT,
                event TEXT,
                evidence_source TEXT,
                person_remarks TEXT,

                FOREIGN KEY(incident_id)
                    REFERENCES incidents(id)
                    ON DELETE CASCADE,

                UNIQUE(incident_id, sequence_no)
            );

            CREATE TABLE IF NOT EXISTS incident_attachments(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                incident_id INTEGER,
                file_path TEXT,
                attachment_type TEXT,
                original_name TEXT,

                FOREIGN KEY(incident_id)
                    REFERENCES incidents(id)
                    ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS audits(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                number TEXT UNIQUE,
                audit_date TEXT,
                audit_type TEXT,
                standard TEXT,
                project TEXT,
                location TEXT,
                department TEXT,
                auditor TEXT,
                lead_auditor TEXT,
                auditee TEXT,
                scope TEXT,
                objective TEXT,
                criteria TEXT,
                start_time TEXT,
                end_time TEXT,
                status TEXT DEFAULT 'Open',
                created_at TEXT
            );

            CREATE TABLE IF NOT EXISTS audit_findings(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                audit_id INTEGER,
                clause TEXT,
                sub_clause TEXT,
                requirement TEXT,
                finding_type TEXT,
                observation TEXT,
                evidence TEXT,
                corrective_action TEXT,
                responsible TEXT,
                target_date TEXT,
                status TEXT DEFAULT 'Open',
                verification TEXT,
                closeout_evidence TEXT,

                FOREIGN KEY(audit_id)
                    REFERENCES audits(id)
                    ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS capa(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                number TEXT UNIQUE,
                source TEXT,
                reference_number TEXT,
                finding TEXT,
                root_cause TEXT,
                corrective_action TEXT,
                preventive_action TEXT,
                responsible TEXT,
                priority TEXT,
                target_date TEXT,
                status TEXT DEFAULT 'Open',
                verification TEXT,
                closeout_evidence TEXT,
                created_at TEXT
            );

            CREATE TABLE IF NOT EXISTS attachments(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                register_type TEXT,
                record_id INTEGER,
                file_path TEXT,
                attachment_type TEXT,
                original_name TEXT,
                created_at TEXT
            );

            CREATE TABLE IF NOT EXISTS employees(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id TEXT,
                name TEXT,
                designation TEXT,
                company TEXT,
                department TEXT,
                contact TEXT,
                active INTEGER DEFAULT 1
            );

            CREATE TABLE IF NOT EXISTS companies(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                company_type TEXT,
                contact_person TEXT,
                contact_number TEXT,
                status TEXT DEFAULT 'Active'
            );

            CREATE TABLE IF NOT EXISTS projects(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project TEXT,
                area TEXT,
                location TEXT,
                description TEXT,
                active INTEGER DEFAULT 1
            );

            CREATE TABLE IF NOT EXISTS categories(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                category TEXT,
                subcategory TEXT,
                active INTEGER DEFAULT 1
            );
            """
        )

        self.conn.commit()

    def execute(self, sql, params=()):
        cursor = self.conn.execute(sql, params)
        self.conn.commit()
        return cursor

    def one(self, sql, params=()):
        return self.conn.execute(sql, params).fetchone()

    def all(self, sql, params=()):
        return self.conn.execute(sql, params).fetchall()

    def setting(self, key, default=""):
        row = self.one(
            "SELECT value FROM settings WHERE key=?",
            (key,)
        )

        if row:
            return row["value"]

        return default

    def set_setting(self, key, value):
        self.execute(
            """
            INSERT INTO settings(key,value)
            VALUES(?,?)
            ON CONFLICT(key)
            DO UPDATE SET value=excluded.value
            """,
            (key, value)
        )


db = DB()


# ============================================================
# NUMBER GENERATOR
# ============================================================

def next_number(prefix, table):

    year = now().year

    row = db.one(
        f"""
        SELECT number
        FROM {table}
        WHERE number LIKE ?
        ORDER BY id DESC
        LIMIT 1
        """,
        (f"{prefix}-{year}-%",)
    )

    number = 1

    if row:

        try:
            number = int(
                row["number"].split("-")[-1]
            ) + 1
        except Exception:
            number = 1

    return f"{prefix}-{year}-{number:05d}"


# ============================================================
# UI HELPERS
# ============================================================

def widget_label(text, size=15, bold=False):

    return Label(
        text=str(text),
        font_size=dp(size),
        bold=bold,
        size_hint_y=None,
        height=dp(38),
        halign="left",
        valign="middle"
    )


class FormScreen(Screen):

    def __init__(self, title, **kwargs):

        super().__init__(**kwargs)

        self.title_text = title

        self.root_box = BoxLayout(
            orientation="vertical",
            padding=dp(8),
            spacing=dp(6)
        )

        self.add_widget(self.root_box)

        self.scroll = ScrollView()

        self.form = GridLayout(
            cols=1,
            spacing=dp(6),
            size_hint_y=None
        )

        self.form.bind(
            minimum_height=self.form.setter("height")
        )

        self.scroll.add_widget(self.form)

        self.root_box.add_widget(
            widget_label(
                title,
                22,
                True
            )
        )

        self.root_box.add_widget(
            self.scroll
        )

    def field(
        self,
        label,
        multiline=False,
        value=""
    ):

        self.form.add_widget(
            widget_label(
                label,
                14,
                True
            )
        )

        height = dp(90 if multiline else 48)

        field = TextInput(
            text=value,
            multiline=multiline,
            size_hint_y=None,
            height=height
        )

        self.form.add_widget(field)

        return field

    def spinner(
        self,
        label,
        values,
        default=None
    ):

        self.form.add_widget(
            widget_label(
                label,
                14,
                True
            )
        )

        spinner = Spinner(
            text=default or values[0],
            values=values,
            size_hint_y=None,
            height=dp(48)
        )

        self.form.add_widget(spinner)

        return spinner

    def button(
        self,
        text,
        callback
    ):

        button = Button(
            text=text,
            size_hint_y=None,
            height=dp(50)
        )

        button.bind(
            on_release=callback
        )

        self.form.add_widget(button)

        return button


# ============================================================
# MAIN APPLICATION
# ============================================================

class HSEApp(App):

    def build(self):

        self.title = APP_NAME

        Window.softinput_mode = "below_target"

        self.sm = ScreenManager()

        self.add_screen(
            "dashboard",
            self.dashboard_screen()
        )

        self.add_screen(
            "inspections",
            self.inspection_screen()
        )

        self.add_screen(
            "incidents",
            self.incident_screen()
        )

        self.add_screen(
            "audits",
            self.audit_screen()
        )

        self.add_screen(
            "capa",
            self.capa_screen()
        )

        self.add_screen(
            "reports",
            self.report_screen()
        )

        self.add_screen(
            "settings",
            self.settings_screen()
        )

        self.show("dashboard")

        return self.sm

    # --------------------------------------------------------
    # SCREEN MANAGEMENT
    # --------------------------------------------------------

    def add_screen(self, name, screen):

        screen.name = name

        self.sm.add_widget(screen)

    def show(self, name):

        self.sm.current = name

    # --------------------------------------------------------
    # COMMON TOP BAR
    # --------------------------------------------------------

    def shell(self, title):

        root = BoxLayout(
            orientation="vertical",
            padding=dp(8),
            spacing=dp(6)
        )

        top = BoxLayout(
            size_hint_y=None,
            height=dp(54),
            spacing=dp(4)
        )

        top.add_widget(
            widget_label(
                title,
                21,
                True
            )
        )

        navigation = [
            ("dashboard", "Home"),
            ("inspections", "Inspections"),
            ("incidents", "Incidents"),
            ("audits", "Audits"),
            ("capa", "CAPA"),
            ("reports", "Reports"),
            ("settings", "Settings")
        ]

        for name, text in navigation:

            button = Button(
                text=text,
                size_hint_x=None,
                width=dp(88)
            )

            button.bind(
                on_release=lambda x, n=name: self.show(n)
            )

            top.add_widget(button)

        root.add_widget(top)

        return root

    # ========================================================
    # DASHBOARD
    # ========================================================

    def dashboard_screen(self):

        screen = Screen()

        root = self.shell(
            "HSE Dashboard"
        )

        scroll = ScrollView()

        box = GridLayout(
            cols=1,
            spacing=dp(8),
            padding=dp(8),
            size_hint_y=None
        )

        box.bind(
            minimum_height=box.setter("height")
        )

        counts = [
            (
                "HSE Inspections",
                db.one(
                    "SELECT COUNT(*) c FROM observations"
                )["c"]
            ),
            (
                "Open Inspections",
                db.one(
                    """
                    SELECT COUNT(*) c
                    FROM observations
                    WHERE status NOT IN ('Closed','Cancelled')
                    """
                )["c"]
            ),
            (
                "Incidents",
                db.one(
                    "SELECT COUNT(*) c FROM incidents"
                )["c"]
            ),
            (
                "Audits",
                db.one(
                    "SELECT COUNT(*) c FROM audits"
                )["c"]
            ),
            (
                "CAPA",
                db.one(
                    "SELECT COUNT(*) c FROM capa"
                )["c"]
            )
        ]

        for title, count in counts:

            box.add_widget(
                widget_label(
                    f"{title}: {count}",
                    18,
                    True
                )
            )

        box.add_widget(
            widget_label(
                "Recent HSE Inspections",
                18,
                True
            )
        )

        recent = db.all(
            """
            SELECT number,obs_type,location,status
            FROM observations
            ORDER BY id DESC
            LIMIT 10
            """
        )

        for row in recent:

            box.add_widget(
                widget_label(
                    f'{row["number"]} | '
                    f'{row["obs_type"]} | '
                    f'{row["location"]} | '
                    f'{row["status"]}',
                    13
                )
            )

        scroll.add_widget(box)

        root.add_widget(scroll)

        screen.add_widget(root)

        return screen

    # ========================================================
    # HSE INSPECTION REGISTER
    # ========================================================

    def inspection_screen(self):

        screen = Screen()

        root = self.shell(
            "HSE Inspection Register"
        )

        button = Button(
            text="+ New Inspection",
            size_hint_y=None,
            height=dp(48)
        )

        button.bind(
            on_release=lambda x: self.new_inspection()
        )

        root.add_widget(button)

        scroll = ScrollView()

        self.obs_box = GridLayout(
            cols=1,
            spacing=dp(5),
            size_hint_y=None
        )

        self.obs_box.bind(
            minimum_height=self.obs_box.setter("height")
        )

        scroll.add_widget(
            self.obs_box
        )

        root.add_widget(scroll)

        screen.add_widget(root)

        Clock.schedule_once(
            lambda dt: self.refresh_observations(),
            0
        )

        return screen

    def refresh_observations(self):

        if not hasattr(self, "obs_box"):
            return

        self.obs_box.clear_widgets()

        rows = db.all(
            """
            SELECT *
            FROM observations
            ORDER BY id DESC
            """
        )

        for row in rows:

            button = Button(
                text=(
                    f'{row["number"]} | '
                    f'{row["obs_date"]} | '
                    f'{row["obs_type"]} | '
                    f'{row["location"]} | '
                    f'{row["priority"]} | '
                    f'{row["status"]}'
                ),
                size_hint_y=None,
                height=dp(58)
            )

            button.bind(
                on_release=lambda x, rec=dict(row):
                    self.view_observation(rec)
            )

            self.obs_box.add_widget(button)

    def new_inspection(self):

        screen = FormScreen(
            "New HSE Inspection"
        )

        screen.name = (
            "inspection_form_"
            + str(id(screen))
        )

        self.sm.add_widget(screen)

        self.sm.current = screen.name

        fields = {}

        names = [
            "Project",
            "Company",
            "Location",
            "Area",
            "Responsible Person",
            "Designation",
            "Employee ID",
            "Observer",
            "Observer Designation",
            "Observer ID",
            "Observer Company"
        ]

        for name in names:

            fields[name] = screen.field(name)

        observation_type = screen.spinner(
            "Observation Type",
            OBS_TYPES
        )

        category = screen.spinner(
            "Category",
            CATEGORIES
        )

        subcategory = screen.field(
            "Subcategory"
        )

        description = screen.field(
            "Observation Description",
            True
        )

        immediate = screen.field(
            "Immediate Action",
            True
        )

        corrective = screen.field(
            "Corrective Action",
            True
        )

        preventive = screen.field(
            "Preventive Action",
            True
        )

        priority = screen.spinner(
            "Priority",
            PRIORITIES
        )

        target = screen.field(
            "Target Completion Date",
            False,
            today()
        )

        status = screen.spinner(
            "Status",
            STATUSES
        )

        closeout = screen.field(
            "Closeout Comments",
            True
        )

        attachments = []

        screen.button(
            "Attach Observation Evidence",
            lambda x:
                self.pick_attachment(
                    attachments,
                    "Observation Evidence"
                )
        )

        screen.button(
            "Attach Closeout Evidence",
            lambda x:
                self.pick_attachment(
                    attachments,
                    "Closeout Evidence"
                )
        )

        def save(_):

            if not fields["Location"].text.strip():

                self.info(
                    "Location is required."
                )

                return

            if not description.text.strip():

                self.info(
                    "Observation Description is required."
                )

                return

            number = next_number(
                "HSE-OBS",
                "observations"
            )

            db.execute(
                """
                INSERT INTO observations(
                    number,
                    obs_date,
                    obs_time,
                    project,
                    company,
                    location,
                    area,
                    responsible,
                    designation,
                    employee_id,
                    observer,
                    observer_designation,
                    observer_id,
                    observer_company,
                    obs_type,
                    category,
                    subcategory,
                    observation,
                    immediate_action,
                    corrective_action,
                    preventive_action,
                    priority,
                    target_date,
                    status,
                    closeout_comments,
                    created_at
                )
                VALUES(
                    ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
                )
                """,
                (
                    number,
                    today(),
                    now().strftime("%H:%M"),
                    fields["Project"].text,
                    fields["Company"].text,
                    fields["Location"].text,
                    fields["Area"].text,
                    fields["Responsible Person"].text,
                    fields["Designation"].text,
                    fields["Employee ID"].text,
                    fields["Observer"].text,
                    fields["Observer Designation"].text,
                    fields["Observer ID"].text,
                    fields["Observer Company"].text,
                    observation_type.text,
                    category.text,
                    subcategory.text,
                    description.text,
                    immediate.text,
                    corrective.text,
                    preventive.text,
                    priority.text,
                    target.text,
                    status.text,
                    closeout.text,
                    now().isoformat()
                )
            )

            record = db.one(
                """
                SELECT id
                FROM observations
                WHERE number=?
                """,
                (number,)
            )

            for path, kind in attachments:

                self.save_attachment(
                    "inspection",
                    record["id"],
                    path,
                    kind
                )

            self.sm.remove_widget(screen)

            self.show("inspections")

            self.refresh_observations()

        screen.button(
            "SAVE INSPECTION",
            save
        )

        screen.button(
            "CANCEL",
            lambda x: (
                self.sm.remove_widget(screen),
                self.show("inspections")
            )
        )

    def view_observation(self, record):

        text = "\n".join(
            [
                f"{key}: {safe(value)}"
                for key, value in record.items()
            ]
        )

        self.info(
            text,
            "Inspection Record"
        )

    # ========================================================
    # INCIDENT REGISTER
    # ========================================================

    def incident_screen(self):

        screen = Screen()

        root = self.shell(
            "Incident Investigation"
        )

        button = Button(
            text="+ New Incident / Investigation",
            size_hint_y=None,
            height=dp(48)
        )

        button.bind(
            on_release=lambda x: self.new_incident()
        )

        root.add_widget(button)

        scroll = ScrollView()

        self.inc_box = GridLayout(
            cols=1,
            spacing=dp(5),
            size_hint_y=None
        )

        self.inc_box.bind(
            minimum_height=self.inc_box.setter("height")
        )

        scroll.add_widget(
            self.inc_box
        )

        root.add_widget(scroll)

        screen.add_widget(root)

        Clock.schedule_once(
            lambda dt: self.refresh_incidents(),
            0
        )

        return screen

    def refresh_incidents(self):

        if not hasattr(self, "inc_box"):
            return

        self.inc_box.clear_widgets()

        rows = db.all(
            """
            SELECT *
            FROM incidents
            ORDER BY id DESC
            """
        )

        for row in rows:

            button = Button(
                text=(
                    f'{row["number"]} | '
                    f'{row["incident_date"]} | '
                    f'{row["incident_type"]} | '
                    f'{row["location"]} | '
                    f'{row["investigation_method"]}'
                ),
                size_hint_y=None,
                height=dp(62)
            )

            button.bind(
                on_release=lambda x, rec=dict(row):
                    self.view_incident(rec)
            )

            self.inc_box.add_widget(button)

    # ========================================================
    # NEW INCIDENT + ICAM TIMELINE
    # ========================================================

    def new_incident(self):

        screen = FormScreen(
            "New Incident Investigation"
        )

        screen.name = (
            "incident_form_"
            + str(id(screen))
        )

        self.sm.add_widget(screen)

        self.sm.current = screen.name

        fields = {}

        names = [
            "Incident Time",
            "Location",
            "Project",
            "Company",
            "Department",
            "Activity",
            "Person Involved",
            "Employee ID",
            "Designation",
            "Supervisor",
            "Witnesses",
            "Equipment"
        ]

        for name in names:

            fields[name] = screen.field(name)

        incident_type = screen.spinner(
            "Incident Type",
            INCIDENT_TYPES
        )

        description = screen.field(
            "Detailed Incident Description",
            True
        )

        immediate = screen.field(
            "Immediate Actions Taken",
            True
        )

        consequences = screen.field(
            "Actual Consequences",
            True
        )

        potential = screen.field(
            "Potential Consequences",
            True
        )

        investigation_method = screen.spinner(
            "Investigation Method",
            INVESTIGATION_METHODS
        )

        status = screen.spinner(
            "Status",
            STATUSES
        )

        # ----------------------------------------------------
        # ICAM TIMELINE
        # ----------------------------------------------------

        screen.form.add_widget(
            widget_label(
                "ICAM / Analysis Timeline",
                18,
                True
            )
        )

        screen.form.add_widget(
            widget_label(
                "Events are stored permanently in SQLite as "
                "Event 1, Event 2, Event 3, etc. "
                "Adding another event does not remove previous events.",
                13
            )
        )

        timeline = GridLayout(
            cols=5,
            spacing=dp(3),
            size_hint_y=None,
            height=dp(48)
        )

        headers = [
            "No.",
            "Date",
            "Time",
            "Event",
            "Evidence / Person / Remarks"
        ]

        for header in headers:

            timeline.add_widget(
                widget_label(
                    header,
                    12,
                    True
                )
            )

        timeline_rows = []

        def add_event(*args):

            # ------------------------------------------------
            # PERMANENT SEQUENTIAL EVENT NUMBER
            # ------------------------------------------------
            sequence = len(timeline_rows) + 1

            cells = []

            values = [
                str(sequence),
                today(),
                "",
                "",
                ""
            ]

            for value in values:

                field = TextInput(
                    text=value,
                    multiline=False,
                    size_hint_y=None,
                    height=dp(48)
                )

                cells.append(field)

                timeline.add_widget(field)

            timeline_rows.append(
                cells
            )

            timeline.height = (
                dp(48)
                * (len(timeline_rows) + 1)
            )

        # Start with Event 1
        add_event()

        screen.form.add_widget(
            timeline
        )

        screen.button(
            "+ ADD TIMELINE EVENT",
            lambda x: add_event()
        )

        # ----------------------------------------------------
        # ICAM ANALYSIS
        # ----------------------------------------------------

        screen.form.add_widget(
            widget_label(
                "ICAM / Root Cause Analysis",
                18,
                True
            )
        )

        analysis = {}

        analysis_names = [
            "Problem / Incident Statement",
            "Immediate / Direct Cause",
            "Contributing Causes",
            "Underlying / System Causes",
            "Failed or Missing Controls",
            "Human / Organizational Factors",
            "Root Cause",
            "Corrective Action",
            "Preventive Action",
            "Lessons Learned"
        ]

        for name in analysis_names:

            analysis[name] = screen.field(
                name,
                True
            )

        attachments = []

        screen.button(
            "Attach Investigation Evidence",
            lambda x:
                self.pick_attachment(
                    attachments,
                    "Investigation Evidence"
                )
        )

        # ----------------------------------------------------
        # SAVE INCIDENT
        # ----------------------------------------------------

        def save(_):

            if not fields["Location"].text.strip():

                self.info(
                    "Location is required."
                )

                return

            if not description.text.strip():

                self.info(
                    "Detailed Incident Description is required."
                )

                return

            number = next_number(
                "HSE-INC",
                "incidents"
            )

            analysis_payload = {
                key: value.text
                for key, value in analysis.items()
            }

            # Save main incident
            db.execute(
                """
                INSERT INTO incidents(
                    number,
                    incident_date,
                    incident_time,
                    location,
                    project,
                    company,
                    department,
                    activity,
                    incident_type,
                    person_involved,
                    employee_id,
                    designation,
                    supervisor,
                    witnesses,
                    description,
                    immediate_action,
                    consequences,
                    potential_consequences,
                    equipment,
                    investigation_method,
                    investigation_details,
                    status,
                    created_at
                )
                VALUES(
                    ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
                )
                """,
                (
                    number,
                    today(),
                    fields["Incident Time"].text,
                    fields["Location"].text,
                    fields["Project"].text,
                    fields["Company"].text,
                    fields["Department"].text,
                    fields["Activity"].text,
                    incident_type.text,
                    fields["Person Involved"].text,
                    fields["Employee ID"].text,
                    fields["Designation"].text,
                    fields["Supervisor"].text,
                    fields["Witnesses"].text,
                    description.text,
                    immediate.text,
                    consequences.text,
                    potential.text,
                    fields["Equipment"].text,
                    investigation_method.text,
                    json.dumps(
                        analysis_payload,
                        ensure_ascii=False
                    ),
                    status.text,
                    now().isoformat()
                )
            )

            incident = db.one(
                """
                SELECT id
                FROM incidents
                WHERE number=?
                """,
                (number,)
            )

            incident_id = incident["id"]

            # ------------------------------------------------
            # SAVE EVERY TIMELINE EVENT INDIVIDUALLY
            # ------------------------------------------------

            for sequence, cells in enumerate(
                timeline_rows,
                start=1
            ):

                db.execute(
                    """
                    INSERT INTO incident_timeline(
                        incident_id,
                        sequence_no,
                        event_date,
                        event_time,
                        event,
                        evidence_source,
                        person_remarks
                    )
                    VALUES(?,?,?,?,?,?,?)
                    """,
                    (
                        incident_id,
                        sequence,
                        cells[1].text,
                        cells[2].text,
                        cells[3].text,
                        cells[4].text,
                        ""
                    )
                )

            # Save attachments
            for path, kind in attachments:

                self.save_attachment(
                    "incident",
                    incident_id,
                    path,
                    kind
                )

            self.sm.remove_widget(
                screen
            )

            self.show(
                "incidents"
            )

            self.refresh_incidents()

            self.info(
                f"Incident {number} saved successfully.\n\n"
                f"{len(timeline_rows)} timeline event(s) saved.\n\n"
                "The timeline events are stored separately "
                "and will remain available."
            )

        screen.button(
            "SAVE INCIDENT / INVESTIGATION",
            save
        )

        screen.button(
            "CANCEL",
            lambda x: (
                self.sm.remove_widget(screen),
                self.show("incidents")
            )
        )

    # ========================================================
    # VIEW INCIDENT
    # ========================================================

    def view_incident(self, record):

        lines = []

        for key, value in record.items():

            if key != "investigation_details":

                lines.append(
                    f"{key}: {safe(value)}"
                )

        lines.append("")
        lines.append(
            "========================================"
        )
        lines.append(
            "ICAM / ANALYSIS TIMELINE"
        )
        lines.append(
            "========================================"
        )

        timeline = db.all(
            """
            SELECT *
            FROM incident_timeline
            WHERE incident_id=?
            ORDER BY sequence_no ASC
            """,
            (record["id"],)
        )

        if not timeline:

            lines.append(
                "No timeline events."
            )

        else:

            for row in timeline:

                lines.append(
                    f'Event {row["sequence_no"]}'
                )

                lines.append(
                    f'Date: {row["event_date"]}'
                )

                lines.append(
                    f'Time: {row["event_time"]}'
                )

                lines.append(
                    f'Event: {row["event"]}'
                )

                lines.append(
                    f'Evidence / Person / Remarks: '
                    f'{row["evidence_source"]}'
                )

                lines.append(
                    ""
                )

        lines.append(
            "========================================"
        )
        lines.append(
            "ICAM ANALYSIS"
        )
        lines.append(
            "========================================"
        )

        details = record.get(
            "investigation_details",
            ""
        )

        if details:

            try:

                payload = json.loads(
                    details
                )

                for key, value in payload.items():

                    lines.append(
                        f"{key}: {safe(value)}"
                    )

            except Exception:

                lines.append(
                    safe(details)
                )

        self.info(
            "\n".join(lines),
            "Incident Investigation"
        )

    # ========================================================
    # AUDIT REGISTER
    # ========================================================

    def audit_screen(self):

        screen = Screen()

        root = self.shell(
            "Audit Register"
        )

        button = Button(
            text="+ New Audit",
            size_hint_y=None,
            height=dp(48)
        )

        button.bind(
            on_release=lambda x: self.new_audit()
        )

        root.add_widget(button)

        scroll = ScrollView()

        box = GridLayout(
            cols=1,
            spacing=dp(5),
            size_hint_y=None
        )

        box.bind(
            minimum_height=box.setter("height")
        )

        rows = db.all(
            """
            SELECT *
            FROM audits
            ORDER BY id DESC
            """
        )

        for row in rows:

            box.add_widget(
                widget_label(
                    f'{row["number"]} | '
                    f'{row["audit_date"]} | '
                    f'{row["standard"]} | '
                    f'{row["audit_type"]} | '
                    f'{row["status"]}',
                    14
                )
            )

        scroll.add_widget(box)

        root.add_widget(scroll)

        screen.add_widget(root)

        return screen

    def new_audit(self):

        screen = FormScreen(
            "New Audit"
        )

        screen.name = (
            "audit_form_"
            + str(id(screen))
        )

        self.sm.add_widget(screen)

        self.sm.current = screen.name

        fields = {}

        names = [
            "Project",
            "Location",
            "Department",
            "Auditor",
            "Lead Auditor",
            "Auditee",
            "Start Time",
            "End Time"
        ]

        for name in names:

            fields[name] = screen.field(
                name
            )

        standard = screen.spinner(
            "Audit Standard",
            [
                "ISO 45001:2018",
                "ISO 14001:2015"
            ]
        )

        audit_type = screen.spinner(
            "Audit Type",
            [
                "Internal Audit",
                "External Audit",
                "Client Audit",
                "Certification Audit",
                "Surveillance Audit",
                "Regulatory Audit",
                "Project Audit",
                "Supplier Audit",
                "Other"
            ]
        )

        scope = screen.field(
            "Scope",
            True
        )

        objective = screen.field(
            "Objective",
            True
        )

        criteria = screen.field(
            "Audit Criteria",
            True
        )

        clause = screen.field(
            "Clause / Sub-Clause"
        )

        requirement = screen.field(
            "Requirement / Reference",
            True
        )

        finding_type = screen.spinner(
            "Finding Type",
            [
                "Positive Observation",
                "Good Practice",
                "Conformity",
                "Opportunity for Improvement",
                "Observation",
                "Minor Nonconformity",
                "Major Nonconformity",
                "Environmental Finding",
                "Legal / Compliance Finding",
                "Other"
            ]
        )

        observation = screen.field(
            "Observation",
            True
        )

        evidence = screen.field(
            "Objective Evidence",
            True
        )

        corrective_action = screen.field(
            "Corrective Action",
            True
        )

        responsible = screen.field(
            "Responsible Person"
        )

        target_date = screen.field(
            "Target Date"
        )

        status = screen.spinner(
            "Status",
            STATUSES
        )

        closeout = screen.field(
            "Closeout Evidence / Verification",
            True
        )

        def save(_):

            number = next_number(
                "HSE-AUD",
                "audits"
            )

            db.execute(
                """
                INSERT INTO audits(
                    number,
                    audit_date,
                    audit_type,
                    standard,
                    project,
                    location,
                    department,
                    auditor,
                    lead_auditor,
                    auditee,
                    scope,
                    objective,
                    criteria,
                    start_time,
                    end_time,
                    status,
                    created_at
                )
                VALUES(
                    ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
                )
                """,
                (
                    number,
                    today(),
                    audit_type.text,
                    standard.text,
                    fields["Project"].text,
                    fields["Location"].text,
                    fields["Department"].text,
                    fields["Auditor"].text,
                    fields["Lead Auditor"].text,
                    fields["Auditee"].text,
                    scope.text,
                    objective.text,
                    criteria.text,
                    fields["Start Time"].text,
                    fields["End Time"].text,
                    status.text,
                    now().isoformat()
                )
            )

            audit = db.one(
                """
                SELECT id
                FROM audits
                WHERE number=?
                """,
                (number,)
            )

            db.execute(
                """
                INSERT INTO audit_findings(
                    audit_id,
                    clause,
                    sub_clause,
                    requirement,
                    finding_type,
                    observation,
                    evidence,
                    corrective_action,
                    responsible,
                    target_date,
                    status,
                    verification,
                    closeout_evidence
                )
                VALUES(
                    ?,?,?,?,?,?,?,?,?,?,?,?,?
                )
                """,
                (
                    audit["id"],
                    clause.text,
                    "",
                    requirement.text,
                    finding_type.text,
                    observation.text,
                    evidence.text,
                    corrective_action.text,
                    responsible.text,
                    target_date.text,
                    status.text,
                    closeout.text,
                    closeout.text
                )
            )

            self.sm.remove_widget(
                screen
            )

            self.show(
                "audits"
            )

            self.info(
                f"Audit {number} saved successfully."
            )

        screen.button(
            "SAVE AUDIT",
            save
        )

        screen.button(
            "CANCEL",
            lambda x: (
                self.sm.remove_widget(screen),
                self.show("audits")
            )
        )

    # ========================================================
    # CAPA
    # ========================================================

    def capa_screen(self):

        screen = Screen()

        root = self.shell(
            "CAPA Register"
        )

        button = Button(
            text="+ New CAPA",
            size_hint_y=None,
            height=dp(48)
        )

        button.bind(
            on_release=lambda x: self.new_capa()
        )

        root.add_widget(button)

        scroll = ScrollView()

        box = GridLayout(
            cols=1,
            spacing=dp(5),
            size_hint_y=None
        )

        box.bind(
            minimum_height=box.setter("height")
        )

        rows = db.all(
            """
            SELECT *
            FROM capa
            ORDER BY id DESC
            """
        )

        for row in rows:

            box.add_widget(
                widget_label(
                    f'{row["number"]} | '
                    f'{row["source"]} | '
                    f'{row["priority"]} | '
                    f'{row["responsible"]} | '
                    f'{row["status"]}',
                    14
                )
            )

        scroll.add_widget(box)

        root.add_widget(scroll)

        screen.add_widget(root)

        return screen

    def new_capa(self):

        screen = FormScreen(
            "New CAPA"
        )

        screen.name = (
            "capa_form_"
            + str(id(screen))
        )

        self.sm.add_widget(screen)

        self.sm.current = screen.name

        source = screen.spinner(
            "Source",
            CAPA_SOURCES
        )

        reference = screen.field(
            "Reference Number"
        )

        finding = screen.field(
            "Finding",
            True
        )

        root_cause = screen.field(
            "Root Cause",
            True
        )

        corrective = screen.field(
            "Corrective Action",
            True
        )

        preventive = screen.field(
            "Preventive Action",
            True
        )

        responsible = screen.field(
            "Responsible Person"
        )

        priority = screen.spinner(
            "Priority",
            PRIORITIES
        )

        target_date = screen.field(
            "Target Date"
        )

        status = screen.spinner(
            "Status",
            STATUSES
        )

        verification = screen.field(
            "Verification",
            True
        )

        closeout = screen.field(
            "Closeout Evidence",
            True
        )

        def save(_):

            number = next_number(
                "HSE-CAPA",
                "capa"
            )

            db.execute(
                """
                INSERT INTO capa(
                    number,
                    source,
                    reference_number,
                    finding,
                    root_cause,
                    corrective_action,
                    preventive_action,
                    responsible,
                    priority,
                    target_date,
                    status,
                    verification,
                    closeout_evidence,
                    created_at
                )
                VALUES(
                    ?,?,?,?,?,?,?,?,?,?,?,?,?,?
                )
                """,
                (
                    number,
                    source.text,
                    reference.text,
                    finding.text,
                    root_cause.text,
                    corrective.text,
                    preventive.text,
                    responsible.text,
                    priority.text,
                    target_date.text,
                    status.text,
                    verification.text,
                    closeout.text,
                    now().isoformat()
                )
            )

            self.sm.remove_widget(
                screen
            )

            self.show(
                "capa"
            )

            self.info(
                f"CAPA {number} saved successfully."
            )

        screen.button(
            "SAVE CAPA",
            save
        )

        screen.button(
            "CANCEL",
            lambda x: (
                self.sm.remove_widget(screen),
                self.show("capa")
            )
        )

    # ========================================================
    # REPORTS
    # ========================================================

    def report_screen(self):

        screen = Screen()

        root = self.shell(
            "Reports & Export"
        )

        box = GridLayout(
            cols=1,
            spacing=dp(8),
            padding=dp(8),
            size_hint_y=None
        )

        box.bind(
            minimum_height=box.setter("height")
        )

        registers = [
            (
                "HSE Inspection Register",
                "observations"
            ),
            (
                "Incident Investigation",
                "incidents"
            ),
            (
                "Audit Register",
                "audits"
            ),
            (
                "CAPA Register",
                "capa"
            )
        ]

        for title, table in registers:

            box.add_widget(
                widget_label(
                    title,
                    18,
                    True
                )
            )

            for extension in [
                "CSV",
                "Excel"
            ]:

                button = Button(
                    text=(
                        f"Export {title} - "
                        f"{extension}"
                    ),
                    size_hint_y=None,
                    height=dp(50)
                )

                button.bind(
                    on_release=lambda x,
                    t=table,
                    e=extension:
                    self.export_register(
                        t,
                        e
                    )
                )

                box.add_widget(button)

        box.add_widget(
            widget_label(
                "CSV contains attachment filenames "
                "because CSV cannot contain binary files. "
                "Excel embeds image attachments when available.",
                13
            )
        )

        backup_button = Button(
            text="CREATE FULL DATABASE BACKUP",
            size_hint_y=None,
            height=dp(50)
        )

        backup_button.bind(
            on_release=lambda x: self.create_backup()
        )

        box.add_widget(
            backup_button
        )

        scroll = ScrollView()

        scroll.add_widget(box)

        root.add_widget(scroll)

        screen.add_widget(root)

        return screen

    # ========================================================
    # SETTINGS
    # ========================================================

    def settings_screen(self):

        screen = FormScreen(
            "Settings"
        )

        # Do NOT add this screen to ScreenManager here.
        # build() adds it once.
        company = screen.field(
            "Company Name",
            False,
            db.setting(
                "company_name",
                ""
            )
        )

        project = screen.field(
            "Project Name",
            False,
            db.setting(
                "project_name",
                ""
            )
        )

        location = screen.field(
            "Default Location",
            False,
            db.setting(
                "default_location",
                ""
            )
        )

        footer = screen.field(
            "Report Footer",
            False,
            db.setting(
                "report_footer",
                ""
            )
        )

        def save(_):

            settings = [
                (
                    "company_name",
                    company.text
                ),
                (
                    "project_name",
                    project.text
                ),
                (
                    "default_location",
                    location.text
                ),
                (
                    "report_footer",
                    footer.text
                )
            ]

            for key, value in settings:

                db.set_setting(
                    key,
                    value
                )

            self.info(
                "Settings saved."
            )

            self.show(
                "dashboard"
            )

        screen.button(
            "SAVE SETTINGS",
            save
        )

        screen.button(
            "BACK",
            lambda x:
            self.show("dashboard")
        )

        return screen

    # ========================================================
    # FILE ATTACHMENTS
    # ========================================================

    def pick_attachment(
        self,
        container,
        kind
    ):

        if filechooser:

            try:

                result = filechooser.open_file(
                    multiple=False
                )

                if result:

                    container.append(
                        (
                            result[0],
                            kind
                        )
                    )

                    self.info(
                        f"Attached: "
                        f"{Path(result[0]).name}"
                    )

                return

            except Exception:
                pass

        chooser = FileChooserListView(
            filters=["*.*"]
        )

        popup = Popup(
            title="Select Attachment",
            content=chooser,
            size_hint=(0.95, 0.9)
        )

        def choose(
            instance,
            selection,
            touch
        ):

            if selection:

                container.append(
                    (
                        selection[0],
                        kind
                    )
                )

                popup.dismiss()

                self.info(
                    f"Attached: "
                    f"{Path(selection[0]).name}"
                )

        chooser.bind(
            on_submit=choose
        )

        popup.open()

    def save_attachment(
        self,
        register_type,
        record_id,
        source,
        kind
    ):

        try:

            source_path = Path(
                source
            )

            if not source_path.exists():
                return

            timestamp = datetime.now().strftime(
                "%Y%m%d%H%M%S"
            )

            destination = (
                ATTACH_DIR
                / (
                    f"{register_type}_"
                    f"{record_id}_"
                    f"{timestamp}_"
                    f"{source_path.name}"
                )
            )

            shutil.copy2(
                source_path,
                destination
            )

            db.execute(
                """
                INSERT INTO attachments(
                    register_type,
                    record_id,
                    file_path,
                    attachment_type,
                    original_name,
                    created_at
                )
                VALUES(?,?,?,?,?,?)
                """,
                (
                    register_type,
                    record_id,
                    str(destination),
                    kind,
                    source_path.name,
                    now().isoformat()
                )
            )

            if register_type == "inspection":

                db.execute(
                    """
                    INSERT INTO observation_attachments(
                        observation_id,
                        file_path,
                        attachment_type,
                        original_name
                    )
                    VALUES(?,?,?,?)
                    """,
                    (
                        record_id,
                        str(destination),
                        kind,
                        source_path.name
                    )
                )

            elif register_type == "incident":

                db.execute(
                    """
                    INSERT INTO incident_attachments(
                        incident_id,
                        file_path,
                        attachment_type,
                        original_name
                    )
                    VALUES(?,?,?,?)
                    """,
                    (
                        record_id,
                        str(destination),
                        kind,
                        source_path.name
                    )
                )

        except Exception as error:

            print(
                "Attachment error:",
                error
            )

    # ========================================================
    # ATTACHMENT INFORMATION
    # ========================================================

    def attachment_text(
        self,
        table,
        record_id
    ):

        if table == "observations":

            register_type = "inspection"

        elif table == "incidents":

            register_type = "incident"

        else:

            register_type = table.rstrip("s")

        rows = db.all(
            """
            SELECT
                original_name,
                file_path,
                attachment_type
            FROM attachments
            WHERE register_type=?
            AND record_id=?
            """,
            (
                register_type,
                record_id
            )
        )

        return " | ".join(
            [
                (
                    f'{row["attachment_type"]}: '
                    f'{row["original_name"]}'
                )
                for row in rows
            ]
        )

    # ========================================================
    # EXPORT
    # ========================================================

    def export_register(
        self,
        table,
        extension
    ):

        rows = db.all(
            f"""
            SELECT *
            FROM {table}
            ORDER BY id DESC
            """
        )

        if not rows:

            self.info(
                "No data to export."
            )

            return

        stamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        # ----------------------------------------------------
        # CSV
        # ----------------------------------------------------

        if extension == "CSV":

            path = (
                EXPORT_DIR
                / f"{table}_{stamp}.csv"
            )

            columns = list(
                rows[0].keys()
            )

            columns.append(
                "Attachments"
            )

            with open(
                path,
                "w",
                newline="",
                encoding="utf-8-sig"
            ) as file:

                writer = csv.writer(
                    file
                )

                writer.writerow(
                    columns
                )

                for row in rows:

                    values = [
                        safe(row[column])
                        for column in
                        rows[0].keys()
                    ]

                    values.append(
                        self.attachment_text(
                            table,
                            row["id"]
                        )
                    )

                    writer.writerow(
                        values
                    )

            self.info(
                "CSV exported to:\n\n"
                f"{path}"
            )

            return

        # ----------------------------------------------------
        # EXCEL
        # ----------------------------------------------------

        if xlsxwriter is None:

            self.info(
                "xlsxwriter is not available."
            )

            return

        path = (
            EXPORT_DIR
            / f"{table}_{stamp}.xlsx"
        )

        workbook = xlsxwriter.Workbook(
            str(path)
        )

        worksheet = workbook.add_worksheet(
            table[:31]
        )

        header_format = workbook.add_format(
            {
                "bold": True,
                "bg_color": "#17365D",
                "font_color": "white",
                "border": 1
            }
        )

        wrap_format = workbook.add_format(
            {
                "text_wrap": True,
                "valign": "top",
                "border": 1
            }
        )

        columns = list(
            rows[0].keys()
        )

        for column_number, header in enumerate(
            columns + ["Attachments"]
        ):

            worksheet.write(
                0,
                column_number,
                header,
                header_format
            )

        for row_number, row in enumerate(
            rows,
            start=1
        ):

            for column_number, column in enumerate(
                columns
            ):

                worksheet.write(
                    row_number,
                    column_number,
                    safe(row[column]),
                    wrap_format
                )

            if table == "observations":

                register_type = "inspection"

            elif table == "incidents":

                register_type = "incident"

            else:

                register_type = table.rstrip("s")

            attachments = db.all(
                """
                SELECT *
                FROM attachments
                WHERE register_type=?
                AND record_id=?
                """,
                (
                    register_type,
                    row["id"]
                )
            )

            attachment_names = " | ".join(
                [
                    safe(
                        attachment["original_name"]
                    )
                    for attachment in attachments
                ]
            )

            worksheet.write(
                row_number,
                len(columns),
                attachment_names,
                wrap_format
            )

            # Embed images where possible
            for attachment in attachments:

                file_path = Path(
                    attachment["file_path"]
                )

                if not file_path.exists():
                    continue

                if file_path.suffix.lower() not in {
                    ".png",
                    ".jpg",
                    ".jpeg",
                    ".bmp"
                }:
                    continue

                try:

                    worksheet.insert_image(
                        row_number,
                        len(columns) + 1,
                        str(file_path),
                        {
                            "x_scale": 0.18,
                            "y_scale": 0.18
                        }
                    )

                except Exception as error:

                    print(
                        "Excel image error:",
                        error
                    )

        worksheet.freeze_panes(
            1,
            0
        )

        worksheet.set_column(
            0,
            len(columns),
            18
        )

        workbook.close()

        self.info(
            "Excel exported to:\n\n"
            f"{path}\n\n"
            "Image attachments are embedded "
            "where supported."
        )

    # ========================================================
    # BACKUP
    # ========================================================

    def backup(self):

        timestamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        path = (
            BACKUP_DIR
            / f"HSE_Backup_{timestamp}.zip"
        )

        with zipfile.ZipFile(
            path,
            "w",
            zipfile.ZIP_DEFLATED
        ) as archive:

            if DB_FILE.exists():

                archive.write(
                    DB_FILE,
                    "database/hse.db"
                )

            for file_path in ATTACH_DIR.rglob("*"):

                if file_path.is_file():

                    archive.write(
                        file_path,
                        f"attachments/{file_path.name}"
                    )

        return path

    def create_backup(self):

        try:

            path = self.backup()

            self.info(
                "Full HSE database backup created:\n\n"
                f"{path}"
            )

        except Exception as error:

            self.info(
                "Backup failed:\n\n"
                f"{error}"
            )

    # ========================================================
    # INFORMATION POPUP
    # ========================================================

    def info(
        self,
        text,
        title="HSE Management System"
    ):

        label = Label(
            text=str(text),
            halign="left",
            valign="top",
            text_size=(dp(500), None)
        )

        popup = Popup(
            title=title,
            content=label,
            size_hint=(0.92, 0.75)
        )

        popup.open()


# ============================================================
# APPLICATION ENTRY POINT
# ============================================================

if __name__ == "__main__":

    HSEApp().run()
