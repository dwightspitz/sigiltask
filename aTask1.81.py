# ============================================================================
# IMPORTS AND DEPENDENCIES
# ============================================================================
# Import all required Python modules for the application
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
import json
import os
import calendar
import uuid

# Try to import drag and drop support
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

# ============================================================================
# MAIN APPLICATION CLASS
# ============================================================================
class ExchangeTaskCreator:
    """
    Main application class that handles the Exchange Task Creator GUI and functionality
    """
    
    # ========================================================================
    # INITIALIZATION AND SETUP
    # ========================================================================
    def __init__(self, root):
        """
        Initialize the application with main window settings and data storage
        """
        # Main window configuration
        self.root = root
        self.root.title("Exchange Task Creator - Enhanced")
        self.root.geometry("1100x680")  # Default full size
        self.root.resizable(False, False)  # Fixed size window
        
        # Apply visual theme styling
        style = ttk.Style()
        style.theme_use('clam')
        
        # Initialize task storage and load existing tasks
        self.tasks = []
        self.categories = ["Work", "Business", "Fun", "Home", "Ideas", "Important", "Needs Preparation", "Personal", "Phone Call", "School", "Sport", "Travel Required"]  # Default categories
        self.locations = []  # Saved locations list
        self.contacts = []  # Contacts with name and email
        self.projects = []  # Projects list for task organization
        self.task_list_visible = True  # Track task list panel visibility
        self.editing_task = None  # Track which task is being edited (None = create mode)
        
        # Window size settings for toggle functionality
        self.full_window_size = "1100x680"  # Full size with task list
        self.compact_window_size = "472x680"  # Compact size without task list
        
        self.load_tasks()
        self.load_categories()  # Load saved categories
        self.load_locations()  # Load saved locations
        self.load_contacts()  # Load saved contacts
        self.load_projects()  # Load saved projects
        
        # Build the user interface
        self.create_widgets()
        
        # Enable drag and drop functionality
        self.setup_drag_and_drop()
    
    # ========================================================================
    # MAIN LAYOUT CREATION
    # ========================================================================
    def create_widgets(self):
        """
        Create the main container and set up the two-panel layout
        """
        # Main container frame that holds everything
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create the left panel (task creation form)
        self.setup_create_task_form()
        
        # Create the right panel (task list tree view)
        self.setup_task_list()
    
    # ========================================================================
    # LEFT PANEL: TASK CREATION FORM
    # ========================================================================
    def setup_create_task_form(self):
        """
        Build the complete task creation form on the left side of the window
        """
        # Main form container
        form_frame = ttk.Frame(self.main_frame)
        form_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # --- BASIC TASK INFORMATION SECTION ---
        task_section = ttk.Frame(form_frame)
        task_section.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5, padx=10)
        
        # Task name input field
        ttk.Label(task_section, text="Task:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.subject_var = tk.StringVar()
        subject_entry = ttk.Entry(task_section, textvariable=self.subject_var, width=40)
        subject_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        
        # Project dropdown with custom entry capability and management button
        ttk.Label(task_section, text="Project:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.project_var = tk.StringVar(value="None")
        project_frame = ttk.Frame(task_section)
        project_frame.grid(row=1, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        
        self.project_combo = ttk.Combobox(project_frame, textvariable=self.project_var, 
                                        values=["None"] + self.projects, width=25)
        self.project_combo.pack(side=tk.LEFT)
        
        # Project management button
        ttk.Button(project_frame, text="📂", width=3, 
                  command=lambda: self.show_project_manager(self.root)).pack(side=tk.LEFT, padx=(2, 0))
        
        # Priority dropdown selection
        ttk.Label(task_section, text="Priority:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.priority_var = tk.StringVar(value="None")
        priority_combo = ttk.Combobox(task_section, textvariable=self.priority_var, 
                                    values=["None", "Low", "Medium", "High"], state="readonly", width=15)
        priority_combo.grid(row=2, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        
        # Status dropdown selection
        ttk.Label(task_section, text="Status:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.status_var = tk.StringVar(value="Not Started")
        status_combo = ttk.Combobox(task_section, textvariable=self.status_var,
                                  values=["Not Started", "In Progress", "Completed", "Waiting", "Deferred"], 
                                  state="readonly", width=15)
        status_combo.grid(row=3, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        
        # Progress section - self-contained frame
        ttk.Label(task_section, text="% Complete:").grid(row=4, column=0, sticky=tk.W, pady=2)
        progress_frame = ttk.Frame(task_section)
        progress_frame.grid(row=4, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        
        # Progress slider and label in contained frame
        self.progress_var = tk.IntVar(value=0)
        progress_scale = ttk.Scale(progress_frame, from_=0, to=100, variable=self.progress_var, 
                                 orient=tk.HORIZONTAL, length=150)
        progress_scale.pack(side=tk.LEFT)
        
        # Fixed-width progress percentage display label
        self.progress_label = ttk.Label(progress_frame, text="0%", width=5, anchor=tk.W)
        self.progress_label.pack(side=tk.LEFT, padx=(10, 0))
        progress_scale.configure(command=self.update_progress_label)
        
        # --- DATES AND TIMING SECTION ---
        dates_section = ttk.Frame(form_frame)
        dates_section.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(15, 5), padx=10)
        
        # Start date dropdown with custom entry capability
        ttk.Label(dates_section, text="Start Date:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.start_date_var = tk.StringVar(value="None")
        self.start_options = {
            "None": None,
            "Today": datetime.now(),
            "Tomorrow": datetime.now() + timedelta(days=1),
            "Yesterday": datetime.now() - timedelta(days=1),
            "Custom": None
        }
        start_combo = ttk.Combobox(dates_section, textvariable=self.start_date_var,
                                 values=list(self.start_options.keys()), state="readonly", width=15)
        start_combo.grid(row=0, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        start_combo.bind("<<ComboboxSelected>>", self.update_start_date)
        
        # Custom start date field (appears when "Custom" is selected)
        self.custom_start_frame = ttk.Frame(dates_section)
        self.custom_start_date_var = tk.StringVar(value=datetime.now().strftime("%b %d, %Y"))
        self.custom_start_entry = ttk.Entry(self.custom_start_frame, textvariable=self.custom_start_date_var, 
                                          width=12, state="readonly")
        self.custom_start_entry.pack(side=tk.LEFT)
        
        custom_start_button = ttk.Button(self.custom_start_frame, text="📅", width=3,
                                       command=lambda: self.show_calendar(self.custom_start_date_var, "Custom Start Date"))
        custom_start_button.pack(side=tk.LEFT, padx=(2, 0))
        
        # Due date smart selection dropdown
        ttk.Label(dates_section, text="Due Date:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.due_date_var = tk.StringVar(value="None")
        self.due_options = {
            "None": None,
            "Today": datetime.now(),
            "Tomorrow": datetime.now() + timedelta(days=1),
            "This Weekend": datetime.now() + timedelta(days=(5-datetime.now().weekday()) % 7),
            "Next Week": datetime.now() + timedelta(days=7),
            "Custom": None
        }
        due_combo = ttk.Combobox(dates_section, textvariable=self.due_date_var,
                               values=list(self.due_options.keys()), state="readonly", width=15)
        due_combo.grid(row=1, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        due_combo.bind("<<ComboboxSelected>>", self.update_due_date)
        
        # Custom due date field (appears when "Custom" is selected)
        self.custom_due_frame = ttk.Frame(dates_section)
        self.custom_due_date_var = tk.StringVar(value=datetime.now().strftime("%b %d, %Y"))
        self.custom_due_entry = ttk.Entry(self.custom_due_frame, textvariable=self.custom_due_date_var, 
                                        width=12, state="readonly")
        self.custom_due_entry.pack(side=tk.LEFT)
        
        custom_due_button = ttk.Button(self.custom_due_frame, text="📅", width=3,
                                     command=lambda: self.show_calendar(self.custom_due_date_var, "Custom Due Date"))
        custom_due_button.pack(side=tk.LEFT, padx=(2, 0))
        
        # --- ENHANCED REMINDER SECTION ---
        reminder_section = ttk.Frame(form_frame)
        reminder_section.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(15, 5), padx=10)
        
        # Enable reminder checkbox
        self.reminder_var = tk.BooleanVar()
        reminder_check = ttk.Checkbutton(reminder_section, text="Set Reminder", variable=self.reminder_var,
                                       command=self.toggle_reminder_options)
        reminder_check.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=2)
        
        # Reminder date field
        ttk.Label(reminder_section, text="Reminder Date:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.reminder_date_var = tk.StringVar(value=datetime.now().strftime("%b %d, %Y"))
        reminder_date_frame = ttk.Frame(reminder_section)
        reminder_date_frame.grid(row=1, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        
        self.reminder_date_entry = ttk.Entry(reminder_date_frame, textvariable=self.reminder_date_var, 
                                           width=12, state="disabled")
        self.reminder_date_entry.pack(side=tk.LEFT)
        
        self.reminder_date_button = ttk.Button(reminder_date_frame, text="📅", width=3, state="disabled",
                                             command=lambda: self.show_calendar(self.reminder_date_var, "Reminder Date"))
        self.reminder_date_button.pack(side=tk.LEFT, padx=(2, 0))
        
        # NEW: Reminder time picker
        ttk.Label(reminder_section, text="Reminder Time:").grid(row=2, column=0, sticky=tk.W, pady=2)
        time_frame = ttk.Frame(reminder_section)
        time_frame.grid(row=2, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        
        # Hour selection
        self.reminder_hour_var = tk.StringVar(value="09")
        hour_spin = ttk.Spinbox(time_frame, from_=0, to=23, width=3, textvariable=self.reminder_hour_var, 
                               format="%02.0f", state="disabled")
        hour_spin.pack(side=tk.LEFT)
        
        ttk.Label(time_frame, text=":").pack(side=tk.LEFT)
        
        # Minute selection
        self.reminder_minute_var = tk.StringVar(value="00")
        minute_spin = ttk.Spinbox(time_frame, from_=0, to=59, width=3, textvariable=self.reminder_minute_var,
                                 format="%02.0f", state="disabled")
        minute_spin.pack(side=tk.LEFT)
        
        # Store references for enabling/disabling
        self.reminder_time_widgets = [hour_spin, minute_spin]
        
        # NEW: Repeat frequency options
        ttk.Label(reminder_section, text="Repeat:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.repeat_frequency_var = tk.StringVar(value="None")
        self.repeat_combo = ttk.Combobox(reminder_section, textvariable=self.repeat_frequency_var,
                                       values=["None", "Daily", "Weekly", "Monthly", "Yearly"], 
                                       state="disabled", width=15)
        self.repeat_combo.grid(row=3, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        self.repeat_combo.bind("<<ComboboxSelected>>", self.update_repeat_options)
        
        # NEW: Repeat end options
        ttk.Label(reminder_section, text="Repeat Until:").grid(row=4, column=0, sticky=tk.W, pady=2)
        self.repeat_end_var = tk.StringVar(value="Never")
        self.repeat_end_combo = ttk.Combobox(reminder_section, textvariable=self.repeat_end_var,
                                           values=["Never", "After X occurrences", "Until date"], 
                                           state="disabled", width=15)
        self.repeat_end_combo.grid(row=4, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        self.repeat_end_combo.bind("<<ComboboxSelected>>", self.update_repeat_end_options)
        
        # Container for repeat end detail options
        self.repeat_end_detail_frame = ttk.Frame(reminder_section)
        self.repeat_end_detail_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=2)
        
        # Repeat count input (for "After X occurrences")
        self.repeat_count_var = tk.StringVar(value="5")
        self.repeat_count_frame = ttk.Frame(self.repeat_end_detail_frame)
        ttk.Label(self.repeat_count_frame, text="Number of occurrences:").pack(side=tk.LEFT)
        self.repeat_count_spin = ttk.Spinbox(self.repeat_count_frame, from_=1, to=365, width=5, 
                                           textvariable=self.repeat_count_var, state="disabled")
        self.repeat_count_spin.pack(side=tk.LEFT, padx=(5, 0))
        
        # Repeat end date input (for "Until date")
        self.repeat_end_date_var = tk.StringVar(value=datetime.now().strftime("%b %d, %Y"))
        self.repeat_end_date_frame = ttk.Frame(self.repeat_end_detail_frame)
        ttk.Label(self.repeat_end_date_frame, text="End date:").pack(side=tk.LEFT)
        self.repeat_end_date_entry = ttk.Entry(self.repeat_end_date_frame, textvariable=self.repeat_end_date_var, 
                                             width=12, state="disabled")
        self.repeat_end_date_entry.pack(side=tk.LEFT, padx=(5, 0))
        self.repeat_end_date_button = ttk.Button(self.repeat_end_date_frame, text="📅", width=3, state="disabled",
                                               command=lambda: self.show_calendar(self.repeat_end_date_var, "Repeat End Date"))
        self.repeat_end_date_button.pack(side=tk.LEFT, padx=(2, 0))
        
        # --- CATEGORIES AND LOCATION SECTION ---
        categories_section = ttk.Frame(form_frame)
        categories_section.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(15, 5), padx=10)
        
        # Category dropdown with custom entry capability - REQUIRED FIELD
        ttk.Label(categories_section, text="Category*:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.category_var = tk.StringVar(value="-- Select Category --")
        self.category_combo = ttk.Combobox(categories_section, textvariable=self.category_var, 
                                         values=["-- Select Category --"] + self.categories, width=20)
        self.category_combo.grid(row=0, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        
        # Owner/Assignee field with contact picker
        ttk.Label(categories_section, text="Owner:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.owner_var = tk.StringVar()
        owner_frame, owner_entry = self.create_contact_autocomplete_entry(categories_section, self.owner_var, width=25)
        owner_frame.grid(row=1, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        
        # Contact management button for easy access
        ttk.Button(categories_section, text="🔧", width=3, 
                  command=lambda: self.show_contact_manager(self.root)).grid(row=1, column=2, padx=(2, 0))
        
        # Location dropdown with custom entry capability
        ttk.Label(categories_section, text="Location:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.location_var = tk.StringVar()
        self.location_combo = ttk.Combobox(categories_section, textvariable=self.location_var, 
                                         values=self.locations, width=20)
        self.location_combo.grid(row=2, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        
        # --- DESCRIPTION SECTION ---
        desc_section = ttk.Frame(form_frame)
        desc_section.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(15, 5), padx=10)
        
        # Description label
        ttk.Label(desc_section, text="Description:", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        # Multi-line text area for detailed description
        self.description_text = tk.Text(desc_section, height=4, width=50, wrap=tk.WORD)
        self.description_text.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # --- ACTION BUTTONS SECTION ---
        buttons_frame = ttk.Frame(form_frame)
        buttons_frame.grid(row=5, column=0, columnspan=2, pady=20)
        
        # Create/Update Task button - dynamic based on editing state
        self.create_update_button = ttk.Button(buttons_frame, text="Create Task", command=self.create_or_update_task)
        self.create_update_button.pack(side=tk.LEFT, padx=5)
        # Clear form button - resets all fields
        ttk.Button(buttons_frame, text="Clear", command=self.clear_form).pack(side=tk.LEFT, padx=5)
        # Toggle task list visibility button
        self.toggle_button = ttk.Button(buttons_frame, text="◀ Hide Tasks", command=self.toggle_task_list)
        self.toggle_button.pack(side=tk.LEFT, padx=(15, 5))
        
        # Configure column weights for proper resizing
        form_frame.columnconfigure(1, weight=1)
        task_section.columnconfigure(1, weight=1)
        dates_section.columnconfigure(1, weight=1)
        categories_section.columnconfigure(1, weight=1)
        desc_section.columnconfigure(0, weight=1)
    
    # ========================================================================
    # RIGHT PANEL: TASK LIST TREE VIEW
    # ========================================================================
    def setup_task_list(self):
        """
        Build the task list tree view on the right side of the window
        """
        # Main container for task list
        self.tasks_frame = ttk.Frame(self.main_frame)
        self.tasks_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0))
        
        # Container for the tree view and scrollbar
        list_frame = ttk.Frame(self.tasks_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tree view widget - shows hierarchical task list with multi-select
        self.task_tree = ttk.Treeview(list_frame, show="tree headings", height=25, selectmode="extended")
        
        # Define columns for the tree view
        self.task_tree["columns"] = ("Priority", "Status", "Due Date", "Owner")
        
        # Configure the main tree column (task names with expand/collapse)
        self.task_tree.heading("#0", text="Task", anchor=tk.W)
        self.task_tree.column("#0", width=200, minwidth=150)
        
        # Configure data columns
        column_widths = {"Priority": 70, "Status": 90, "Due Date": 90, "Owner": 120}
        for col in self.task_tree["columns"]:
            self.task_tree.heading(col, text=col)
            self.task_tree.column(col, width=column_widths.get(col, 70), minwidth=60)
        
        # Vertical scrollbar for long task lists
        tree_scroll_y = ttk.Scrollbar(list_frame, orient="vertical", command=self.task_tree.yview)
        self.task_tree.configure(yscrollcommand=tree_scroll_y.set)
        
        # Layout the tree view and scrollbar
        self.task_tree.pack(side="left", fill="both", expand=True)
        tree_scroll_y.pack(side="right", fill="y")
        
        # --- TASK MANAGEMENT BUTTONS ---
        task_buttons_frame = ttk.Frame(self.tasks_frame)
        task_buttons_frame.pack(fill=tk.X, pady=(5, 10), padx=10)
        
        # Button order: Refresh, Edit, Delete, Import, Export, Create Event
        ttk.Button(task_buttons_frame, text="Refresh", command=self.refresh_task_list, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(task_buttons_frame, text="Edit", command=self.edit_selected_task, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(task_buttons_frame, text="Delete", command=self.delete_selected_task, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(task_buttons_frame, text="Import", command=self.manual_import, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(task_buttons_frame, text="Export", command=self.export_selected_task, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(task_buttons_frame, text="Create Event", command=self.create_calendar_event, width=14).pack(side=tk.LEFT, padx=2)
        
        # Double-click event to view detailed task information
        self.task_tree.bind("<Double-1>", self.view_task_details)
        
        # Load and display existing tasks
        self.refresh_task_list()
    
    # ========================================================================
    # NEW: ENHANCED REMINDER FUNCTIONALITY
    # ========================================================================
    def toggle_reminder_options(self):
        """
        Enable/disable reminder options based on checkbox state
        """
        state = "normal" if self.reminder_var.get() else "disabled"
        
        # Toggle reminder date controls
        self.reminder_date_entry.config(state=state)
        self.reminder_date_button.config(state=state)
        
        # Toggle time controls
        for widget in self.reminder_time_widgets:
            widget.config(state=state)
        
        # Toggle repeat controls
        self.repeat_combo.config(state=state)
        self.repeat_end_combo.config(state=state)
        
        # If disabling, also hide any repeat end detail options
        if not self.reminder_var.get():
            self.hide_all_repeat_end_options()
    
    def update_repeat_options(self, event=None):
        """
        Handle changes to repeat frequency selection
        """
        frequency = self.repeat_frequency_var.get()
        if frequency == "None":
            # Hide repeat end options when no repeat is selected
            self.repeat_end_combo.config(state="disabled")
            self.hide_all_repeat_end_options()
        else:
            # Enable repeat end options when a frequency is selected
            self.repeat_end_combo.config(state="normal")
    
    def update_repeat_end_options(self, event=None):
        """
        Show/hide repeat end detail options based on selection
        """
        end_type = self.repeat_end_var.get()
        
        # Hide all options first
        self.hide_all_repeat_end_options()
        
        # Show appropriate option based on selection
        if end_type == "After X occurrences":
            self.repeat_count_frame.pack(fill=tk.X, pady=2)
            self.repeat_count_spin.config(state="normal")
        elif end_type == "Until date":
            self.repeat_end_date_frame.pack(fill=tk.X, pady=2)
            self.repeat_end_date_entry.config(state="readonly")
            self.repeat_end_date_button.config(state="normal")
    
    def hide_all_repeat_end_options(self):
        """
        Hide all repeat end detail option frames
        """
        self.repeat_count_frame.pack_forget()
        self.repeat_end_date_frame.pack_forget()
        
        # Disable controls
        self.repeat_count_spin.config(state="disabled")
        self.repeat_end_date_entry.config(state="disabled")
        self.repeat_end_date_button.config(state="disabled")
    
    def get_reminder_time_string(self):
        """
        Format the reminder time as a string
        """
        if not self.reminder_var.get():
            return ""
        
        hour = self.reminder_hour_var.get().zfill(2)
        minute = self.reminder_minute_var.get().zfill(2)
        return f"{hour}:{minute}"
    
    def get_repeat_settings(self):
        """
        Get all repeat settings as a dictionary
        """
        if not self.reminder_var.get() or self.repeat_frequency_var.get() == "None":
            return None
        
        settings = {
            "frequency": self.repeat_frequency_var.get(),
            "end_type": self.repeat_end_var.get()
        }
        
        if self.repeat_end_var.get() == "After X occurrences":
            settings["count"] = int(self.repeat_count_var.get())
        elif self.repeat_end_var.get() == "Until date":
            settings["end_date"] = self.convert_display_to_iso(self.repeat_end_date_var.get())
        
        return settings
    
    # ========================================================================
    # DRAG AND DROP FUNCTIONALITY
    # ========================================================================
    def setup_drag_and_drop(self):
        """
        Enable drag and drop functionality for importing tasks from files
        """
        try:
            # Try to enable drag and drop (requires tkinterdnd2)
            self.root.drop_target_register('DND_Files')
            self.root.dnd_bind('<<DropEnter>>', self.on_drag_enter)
            self.root.dnd_bind('<<DropLeave>>', self.on_drag_leave)
            self.root.dnd_bind('<<Drop>>', self.on_file_drop)
        except:
            # Drag and drop not available, but Import button handles manual import
            pass
    
    def manual_import(self):
        """
        Manual file import dialog as fallback for drag and drop
        """
        files = filedialog.askopenfilenames(
            title="Select files to import",
            filetypes=[
                ("All supported", "*.ics;*.json;*.txt;*.csv"),
                ("iCalendar files", "*.ics"),
                ("JSON files", "*.json"),
                ("Text files", "*.txt"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        
        if files:
            imported_count = 0
            for file_path in files:
                if self.import_tasks_from_file(file_path):
                    imported_count += 1
            
            if imported_count > 0:
                messagebox.showinfo("Import Success", 
                                  f"Successfully imported tasks from {imported_count} file(s)!")
                self.refresh_task_list()
            else:
                messagebox.showwarning("Import Failed", 
                                     "No tasks could be imported from the selected files.")
    
    def on_drag_enter(self, event):
        """
        Visual feedback when files are dragged over the window
        """
        self.root.config(relief='ridge', borderwidth=2)
        return 'copy'
    
    def on_drag_leave(self, event):
        """
        Remove visual feedback when drag leaves the window
        """
        self.root.config(relief='flat', borderwidth=0)
    
    def on_file_drop(self, event):
        """
        Handle files dropped onto the window
        """
        self.root.config(relief='flat', borderwidth=0)
        
        # Get list of dropped files
        files = event.data
        if isinstance(files, str):
            files = [files]
        
        imported_count = 0
        for file_path in files:
            if self.import_tasks_from_file(file_path):
                imported_count += 1
        
        if imported_count > 0:
            messagebox.showinfo("Import Success", 
                              f"Successfully imported tasks from {imported_count} file(s)!")
            self.refresh_task_list()
        else:
            messagebox.showwarning("Import Failed", 
                                 "No tasks could be imported from the dropped files.")
    
    def import_tasks_from_file(self, file_path):
        """
        Import tasks from various file formats
        """
        try:
            file_ext = file_path.lower().split('.')[-1]
            
            if file_ext == 'ics':
                return self.import_from_ics(file_path)
            elif file_ext == 'json':
                return self.import_from_json(file_path)
            elif file_ext in ['txt', 'csv']:
                return self.import_from_text(file_path)
            else:
                return False
        except Exception as e:
            print(f"Error importing {file_path}: {e}")
            return False
    
    def import_from_ics(self, file_path):
        """
        Import tasks from iCalendar (.ics) files - supports VTODO and VEVENT
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Parse both VTODO (tasks) and VEVENT (calendar events that can become tasks)
            items = []
            current_item = {}
            in_item = False
            item_type = None
            
            for line in content.split('\n'):
                line = line.strip()
                
                if line == 'BEGIN:VTODO':
                    in_item = True
                    item_type = 'VTODO'
                    current_item = {'type': 'VTODO'}
                elif line == 'BEGIN:VEVENT':
                    in_item = True
                    item_type = 'VEVENT'
                    current_item = {'type': 'VEVENT'}
                elif line in ['END:VTODO', 'END:VEVENT']:
                    if current_item:
                        items.append(current_item)
                    in_item = False
                    item_type = None
                elif in_item and ':' in line:
                    # Handle line folding (long lines split across multiple lines)
                    if line.startswith(' ') or line.startswith('\t'):
                        # This is a continuation of the previous line
                        continue
                    
                    key, value = line.split(':', 1)
                    # Handle parameters (e.g., DTSTART;VALUE=DATE:20241215)
                    if ';' in key:
                        key = key.split(';')[0]
                    current_item[key] = value
            
            # Convert to our task format
            imported_count = 0
            for item in items:
                # Create task from either VTODO or VEVENT
                subject = item.get('SUMMARY', 'Imported Task')
                
                # For events, add indication that it was converted from an event
                if item.get('type') == 'VEVENT':
                    subject = f"📅 {subject}"
                
                task = {
                    "id": len(self.tasks) + 1,
                    "subject": subject,
                    "priority": self.convert_ics_priority(item.get('PRIORITY', '0')),
                    "status": self.convert_ics_status(item.get('STATUS', 'NEEDS-ACTION')),
                    "progress": int(item.get('PERCENT-COMPLETE', '0')),
                    "start_date": self.convert_ics_date(item.get('DTSTART', '')),
                    "due_date": self.convert_ics_date(item.get('DUE', '') or item.get('DTEND', '')),
                    "due_date_type": "Custom",
                    "reminder": False,
                    "reminder_date": "",
                    "reminder_time": "",
                    "repeat_settings": None,
                    "category": item.get('CATEGORIES', ''),
                    "owner": item.get('ORGANIZER', '').replace('mailto:', '') if item.get('ORGANIZER') else "",
                    "location": item.get('LOCATION', ''),
                    "description": item.get('DESCRIPTION', '').replace('\\n', '\n').replace('\\,', ',').replace('\\;', ';'),
                    "created_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                self.tasks.append(task)
                imported_count += 1
            
            return imported_count > 0
        except Exception as e:
            print(f"Error importing ICS file: {e}")
            return False
    
    def import_from_json(self, file_path):
        """
        Import tasks from JSON files
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            if isinstance(data, list):
                # Assume it's a list of tasks
                for task_data in data:
                    if isinstance(task_data, dict) and 'subject' in task_data:
                        task = {
                            "id": len(self.tasks) + 1,
                            "subject": task_data.get('subject', 'Imported Task'),
                            "priority": task_data.get('priority', 'None'),
                            "status": task_data.get('status', 'Not Started'),
                            "progress": task_data.get('progress', 0),
                            "start_date": task_data.get('start_date', ''),
                            "due_date": task_data.get('due_date', ''),
                            "due_date_type": task_data.get('due_date_type', 'Custom'),
                            "reminder": task_data.get('reminder', False),
                            "reminder_date": task_data.get('reminder_date', ''),
                            "reminder_time": task_data.get('reminder_time', ''),
                            "repeat_settings": task_data.get('repeat_settings'),
                            "category": task_data.get('category', ''),
                            "owner": task_data.get('owner', ''),
                            "location": task_data.get('location', ''),
                            "description": task_data.get('description', ''),
                            "created_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        }
                        self.tasks.append(task)
                return len(data) > 0
            return False
        except Exception as e:
            return False
    
    def import_from_text(self, file_path):
        """
        Import tasks from text/CSV files (simple format)
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            for line in lines:
                line = line.strip()
                if line and not line.startswith('#'):  # Skip empty lines and comments
                    # Simple text import - each line becomes a task
                    task = {
                        "id": len(self.tasks) + 1,
                        "subject": line,
                        "priority": "None",
                        "status": "Not Started",
                        "progress": 0,
                        "start_date": "",
                        "due_date": "",
                        "due_date_type": "Custom",
                        "reminder": False,
                        "reminder_date": "",
                        "reminder_time": "",
                        "repeat_settings": None,
                        "category": "",
                        "owner": "",
                        "location": "",
                        "description": "",
                        "created_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                    self.tasks.append(task)
            
            return len([l for l in lines if l.strip() and not l.strip().startswith('#')]) > 0
        except Exception as e:
            return False
    
    def convert_ics_priority(self, priority_str):
        """Convert ICS priority (1-9) to our priority system"""
        try:
            priority = int(priority_str)
            if priority <= 3:
                return "High"
            elif priority <= 6:
                return "Medium"
            elif priority <= 9:
                return "Low"
        except:
            pass
        return "None"
    
    def convert_ics_status(self, status_str):
        """Convert ICS status to our status system"""
        status_map = {
            "NEEDS-ACTION": "Not Started",
            "IN-PROCESS": "In Progress",
            "COMPLETED": "Completed",
            "CANCELLED": "Deferred"
        }
        return status_map.get(status_str, "Not Started")
    
    def convert_ics_date(self, date_str):
        """Convert ICS date format to our ISO format"""
        if not date_str:
            return ""
        
        try:
            # Handle different ICS date formats
            if 'VALUE=DATE:' in date_str:
                date_str = date_str.split('VALUE=DATE:')[1]
            elif ':' in date_str:
                date_str = date_str.split(':')[1]
            
            # Remove time portion if present
            if 'T' in date_str:
                date_str = date_str.split('T')[0]
            
            # Convert YYYYMMDD to YYYY-MM-DD
            if len(date_str) == 8 and date_str.isdigit():
                return f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}"
        except:
            pass
        
        return ""
    
    # ========================================================================
    # CALENDAR EVENT CREATION
    # ========================================================================
    def create_calendar_event(self):
        """
        Convert selected task(s) to calendar event(s)
        """
        selected_items = self.task_tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "Please select a task to convert to calendar event.")
            return
        
        # Filter out categories and get only tasks
        selected_tasks = []
        for item_id in selected_items:
            item = self.task_tree.item(item_id)
            if item["tags"]:
                task_id = int(item["tags"][0])
                task = next((t for t in self.tasks if t["id"] == task_id), None)
                if task:
                    selected_tasks.append(task)
        
        if not selected_tasks:
            messagebox.showwarning("Warning", "Please select actual tasks, not categories.")
            return
        
        # Show event creation dialog
        self.show_event_creation_dialog(selected_tasks)
    
    def show_event_creation_dialog(self, tasks):
        """
        Show dialog to configure calendar event details - streamlined layout
        """
        # Create event configuration window
        event_window = tk.Toplevel(self.root)
        event_window.title("Create Calendar Event")
        event_window.geometry("500x650")  # Extended height
        event_window.resizable(False, False)
        event_window.grab_set()  # Make window modal
        
        # Main frame (no scrolling)
        main_frame = ttk.Frame(event_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # === BASIC EVENT INFORMATION ===
        basic_frame = ttk.LabelFrame(main_frame, text="Event Details", padding="8")
        basic_frame.pack(fill=tk.X, pady=(0, 8))
        
        # Event title/summary
        ttk.Label(basic_frame, text="Subject:").grid(row=0, column=0, sticky=tk.W, pady=2)
        if len(tasks) == 1:
            task_name = self.format_task_name_with_project(tasks[0]['subject'], tasks[0].get('project', ''))
            event_subject_var = tk.StringVar(value=task_name)
        else:
            event_subject_var = tk.StringVar(value=f"Work on {len(tasks)} tasks")
        event_subject_entry = ttk.Entry(basic_frame, textvariable=event_subject_var, width=45)
        event_subject_entry.grid(row=0, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        
        # Location
        ttk.Label(basic_frame, text="Location:").grid(row=1, column=0, sticky=tk.W, pady=2)
        event_location_var = tk.StringVar(value=tasks[0].get('location', '') if len(tasks) == 1 else '')
        event_location_entry = ttk.Entry(basic_frame, textvariable=event_location_var, width=45)
        event_location_entry.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        
        # Description (smaller)
        ttk.Label(basic_frame, text="Description:").grid(row=2, column=0, sticky=tk.NW, pady=(5, 2))
        event_description_text = tk.Text(basic_frame, height=3, width=45, wrap=tk.WORD)
        event_description_text.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 2), padx=(5, 0))
        
        # Pre-fill description with task details
        if len(tasks) == 1:
            desc_parts = []
            if tasks[0].get('description'):
                desc_parts.append(tasks[0]['description'])
            desc_parts.append(f"Priority: {tasks[0].get('priority', 'None')} | Status: {tasks[0].get('status', 'Not Started')} ({tasks[0].get('progress', 0)}%)")
            event_description_text.insert("1.0", "\n".join(desc_parts))
        
        # === DATE, TIME & SETTINGS COMBINED ===
        datetime_frame = ttk.LabelFrame(main_frame, text="Date, Time & Settings", padding="8")
        datetime_frame.pack(fill=tk.X, pady=(0, 8))
        
        # Event date and all-day option on same row
        ttk.Label(datetime_frame, text="Date:").grid(row=0, column=0, sticky=tk.W, pady=2)
        event_date_var = tk.StringVar(value=datetime.now().strftime("%b %d, %Y"))
        event_date_frame = ttk.Frame(datetime_frame)
        event_date_frame.grid(row=0, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        
        event_date_entry = ttk.Entry(event_date_frame, textvariable=event_date_var, width=12, state="readonly")
        event_date_entry.pack(side=tk.LEFT)
        
        event_date_button = ttk.Button(event_date_frame, text="📅", width=3,
                                     command=lambda: self.show_calendar(event_date_var, "Event Date"))
        event_date_button.pack(side=tk.LEFT, padx=(2, 0))
        
        # All-day event option on same row
        all_day_var = tk.BooleanVar()
        all_day_check = ttk.Checkbutton(datetime_frame, text="All-day event", variable=all_day_var)
        all_day_check.grid(row=0, column=2, sticky=tk.W, pady=2, padx=(20, 0))
        
        # Start and End time on same row
        ttk.Label(datetime_frame, text="Time:").grid(row=1, column=0, sticky=tk.W, pady=2)
        time_frame = ttk.Frame(datetime_frame)
        time_frame.grid(row=1, column=1, columnspan=2, sticky=tk.W, pady=2, padx=(5, 0))
        
        # Start time
        event_start_hour_var = tk.StringVar(value="09")
        start_hour_spin = ttk.Spinbox(time_frame, from_=0, to=23, width=3, textvariable=event_start_hour_var, format="%02.0f")
        start_hour_spin.pack(side=tk.LEFT)
        
        ttk.Label(time_frame, text=":").pack(side=tk.LEFT)
        
        event_start_minute_var = tk.StringVar(value="30")
        start_minute_spin = ttk.Spinbox(time_frame, from_=0, to=59, width=3, textvariable=event_start_minute_var, format="%02.0f")
        start_minute_spin.pack(side=tk.LEFT)
        
        ttk.Label(time_frame, text=" to ").pack(side=tk.LEFT, padx=(10, 10))
        
        # End time
        event_end_hour_var = tk.StringVar(value="11")
        end_hour_spin = ttk.Spinbox(time_frame, from_=0, to=23, width=3, textvariable=event_end_hour_var, format="%02.0f")
        end_hour_spin.pack(side=tk.LEFT)
        
        ttk.Label(time_frame, text=":").pack(side=tk.LEFT)
        
        event_end_minute_var = tk.StringVar(value="00")
        end_minute_spin = ttk.Spinbox(time_frame, from_=0, to=59, width=3, textvariable=event_end_minute_var, format="%02.0f")
        end_minute_spin.pack(side=tk.LEFT)
        
        # Store time controls for toggling
        time_controls = [start_hour_spin, start_minute_spin, end_hour_spin, end_minute_spin]
        all_day_check.configure(command=lambda: self.toggle_time_controls(all_day_var.get(), time_controls))
        
        # Priority and Show as on same row
        ttk.Label(datetime_frame, text="Priority:").grid(row=2, column=0, sticky=tk.W, pady=2)
        priority_var = tk.StringVar(value=tasks[0].get('priority', 'Medium') if len(tasks) == 1 else 'Medium')
        priority_combo = ttk.Combobox(datetime_frame, textvariable=priority_var, 
                                    values=["High", "Medium", "Low"], state="readonly", width=10)
        priority_combo.grid(row=2, column=1, sticky=tk.W, pady=2, padx=(5, 0))
        
        ttk.Label(datetime_frame, text="Show as:").grid(row=2, column=2, sticky=tk.W, pady=2, padx=(20, 0))
        show_as_var = tk.StringVar(value="Busy")
        show_as_combo = ttk.Combobox(datetime_frame, textvariable=show_as_var,
                                   values=["Busy", "Free", "Tentative", "Out of Office"], state="readonly", width=12)
        show_as_combo.grid(row=2, column=3, sticky=tk.W, pady=2, padx=(5, 0))
        
        # === ORGANIZER & ATTENDEES ===
        people_frame = ttk.LabelFrame(main_frame, text="Organizer & Attendees", padding="8")
        people_frame.pack(fill=tk.X, pady=(0, 8))
        
        # Organizer
        ttk.Label(people_frame, text="Organizer:").grid(row=0, column=0, sticky=tk.W, pady=2)
        organizer_var = tk.StringVar(value="Randy Buhr <rabuhr@fseinc.net>")
        organizer_frame, organizer_entry = self.create_contact_autocomplete_entry(people_frame, organizer_var, width=30)
        organizer_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        
        # Contact management button
        ttk.Button(people_frame, text="Contacts", 
                  command=lambda: self.show_contact_manager(event_window)).grid(row=0, column=2, padx=(10, 0))
        
        # Attendees (compact)
        ttk.Label(people_frame, text="Attendees:").grid(row=1, column=0, sticky=tk.NW, pady=(5, 2))
        
        # Attendees frame with add functionality
        attendees_frame = ttk.Frame(people_frame)
        attendees_frame.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 2), padx=(5, 0))
        
        # Attendees listbox (smaller)
        attendees_listbox = tk.Listbox(attendees_frame, height=3)
        attendees_listbox.pack(fill=tk.X, pady=(0, 5))
        
        # Add attendee controls
        add_attendee_frame = ttk.Frame(attendees_frame)
        add_attendee_frame.pack(fill=tk.X)
        
        attendee_var = tk.StringVar()
        attendee_entry_frame, attendee_entry = self.create_contact_autocomplete_entry(add_attendee_frame, attendee_var, width=20)
        attendee_entry_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        def add_attendee():
            attendee = attendee_var.get().strip()
            if attendee:
                # Check if it's already in the list
                current_attendees = [attendees_listbox.get(i) for i in range(attendees_listbox.size())]
                if attendee not in current_attendees:
                    attendees_listbox.insert(tk.END, attendee)
                    attendee_var.set("")
                else:
                    messagebox.showwarning("Duplicate", "This attendee is already in the list.")
        
        def remove_attendee():
            selection = attendees_listbox.curselection()
            if selection:
                attendees_listbox.delete(selection[0])
        
        ttk.Button(add_attendee_frame, text="Add", command=add_attendee).pack(side=tk.LEFT, padx=(5, 2))
        ttk.Button(add_attendee_frame, text="Remove", command=remove_attendee).pack(side=tk.LEFT, padx=2)
        
        # Bind Enter key to add attendee
        attendee_entry.bind('<Return>', lambda e: add_attendee())
        
        # === REMINDER & OPTIONS ===
        reminder_frame = ttk.LabelFrame(main_frame, text="Reminder & Options", padding="8")
        reminder_frame.pack(fill=tk.X, pady=(0, 8))
        
        # Reminder enabled and time on same row
        reminder_enabled_var = tk.BooleanVar(value=True)
        reminder_check = ttk.Checkbutton(reminder_frame, text="Set reminder", variable=reminder_enabled_var)
        reminder_check.grid(row=0, column=0, sticky=tk.W, pady=2)
        
        reminder_time_var = tk.StringVar(value="30 minutes before")
        reminder_combo = ttk.Combobox(reminder_frame, textvariable=reminder_time_var, width=18,
                                    values=["15 minutes before", "30 minutes before", "1 hour before", 
                                           "2 hours before", "1 day before", "1 week before"])
        reminder_combo.grid(row=0, column=1, sticky=tk.W, pady=2, padx=(10, 0))
        
        # Options on same row
        private_var = tk.BooleanVar()
        private_check = ttk.Checkbutton(reminder_frame, text="Private", variable=private_var)
        private_check.grid(row=1, column=0, sticky=tk.W, pady=2)
        
        allow_counter_var = tk.BooleanVar(value=True)
        counter_check = ttk.Checkbutton(reminder_frame, text="Allow counter proposals", variable=allow_counter_var)
        counter_check.grid(row=1, column=1, sticky=tk.W, pady=2, padx=(10, 0))
        
        # === BUTTONS ===
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=(15, 0))
        
        ttk.Button(buttons_frame, text="Create Event File", 
                  command=lambda: self.generate_professional_calendar_event(
                      tasks, event_subject_var.get(), event_description_text.get("1.0", tk.END).strip(),
                      event_location_var.get(), event_date_var.get(), 
                      event_start_hour_var.get(), event_start_minute_var.get(),
                      event_end_hour_var.get(), event_end_minute_var.get(),
                      all_day_var.get(), organizer_var.get(), self.get_attendees_list(attendees_listbox),
                      priority_var.get(), show_as_var.get(), private_var.get(), allow_counter_var.get(),
                      reminder_enabled_var.get(), reminder_time_var.get(), event_window)).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(buttons_frame, text="Cancel", command=event_window.destroy).pack(side=tk.RIGHT, padx=5)
        
        # Configure grid weights
        basic_frame.columnconfigure(1, weight=1)
        people_frame.columnconfigure(1, weight=1)
        
        # Center the window
        event_window.update_idletasks()
        x = (event_window.winfo_screenwidth() // 2) - (event_window.winfo_width() // 2)
        y = (event_window.winfo_screenheight() // 2) - (event_window.winfo_height() // 2)
        event_window.geometry(f"+{x}+{y}")
    
    def generate_professional_calendar_event(self, tasks, subject, description, location, event_date, 
                                           start_hour, start_minute, end_hour, end_minute, all_day,
                                           organizer, attendees, priority, show_as, private, allow_counter,
                                           reminder_enabled, reminder_time, dialog_window):
        """
        Generate professional calendar event .ics file matching EM Client format
        """
        try:
            # Create filename
            safe_filename = "".join(c for c in subject if c.isalnum() or c in (' ', '-', '_')).rstrip()
            default_name = f"Event_{safe_filename}.ics"
            
            # Show file save dialog
            filename = filedialog.asksaveasfilename(
                defaultextension=".ics",
                initialfile=default_name,
                filetypes=[("iCalendar files", "*.ics"), ("All files", "*.*")]
            )
            
            if filename:
                self.create_professional_ics_file(tasks, filename, subject, description, location, event_date,
                                                start_hour, start_minute, end_hour, end_minute, all_day,
                                                organizer, attendees, priority, show_as, private, allow_counter,
                                                reminder_enabled, reminder_time)
                dialog_window.destroy()
                messagebox.showinfo("Success", f"Professional calendar event created: {filename}\n\nThis file is compatible with EM Client, Outlook, and other calendar applications!")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create calendar event: {e}")
    
    def create_professional_ics_file(self, tasks, filename, subject, description, location, event_date,
                                   start_hour, start_minute, end_hour, end_minute, all_day,
                                   organizer, attendees, priority, show_as, private, allow_counter,
                                   reminder_enabled, reminder_time):
        """
        Create professional iCalendar (.ics) file matching EM Client format
        """
        ics_content = [
            "BEGIN:VCALENDAR",
            "VERSION:2.0",
            "PRODID:-//Exchange Task Creator//Professional Event Generator//EN",
            "BEGIN:VTIMEZONE",
            "TZID:Eastern Standard Time",
            "BEGIN:STANDARD",
            "DTSTART:16011101T020000",
            "TZOFFSETFROM:-0400",
            "TZOFFSETTO:-0500",
            "RRULE:FREQ=YEARLY;BYDAY=1SU;BYMONTH=11",
            "TZNAME:Standard Time",
            "END:STANDARD",
            "BEGIN:DAYLIGHT",
            "DTSTART:16010302T020000",
            "TZOFFSETFROM:-0500",
            "TZOFFSETTO:-0400",
            "RRULE:FREQ=YEARLY;BYDAY=2SU;BYMONTH=3",
            "TZNAME:Daylight Savings Time",
            "END:DAYLIGHT",
            "END:VTIMEZONE"
        ]
        
        # Convert event date
        event_date_obj = datetime.strptime(event_date, "%b %d, %Y")
        
        # Create unique UID
        uid = str(uuid.uuid4())
        
        # Calculate event times
        if all_day:
            dtstart = event_date_obj.strftime('%Y%m%d')
            dtend = (event_date_obj + timedelta(days=1)).strftime('%Y%m%d')
            dtstart_line = f"DTSTART;VALUE=DATE:{dtstart}"
            dtend_line = f"DTEND;VALUE=DATE:{dtend}"
        else:
            start_time = event_date_obj.replace(hour=int(start_hour), minute=int(start_minute))
            end_time = event_date_obj.replace(hour=int(end_hour), minute=int(end_minute))
            
            dtstart = start_time.strftime('%Y%m%dT%H%M%S')
            dtend = end_time.strftime('%Y%m%dT%H%M%S')
            dtstart_line = f"DTSTART;TZID=Eastern Standard Time:{dtstart}"
            dtend_line = f"DTEND;TZID=Eastern Standard Time:{dtend}"
        
        # Map priority
        priority_map = {"High": "1", "Medium": "5", "Low": "9"}
        ics_priority = priority_map.get(priority, "5")
        
        # Map show as status
        status_map = {"Busy": "BUSY", "Free": "FREE", "Tentative": "TENTATIVE", "Out of Office": "OOF"}
        busy_status = status_map.get(show_as, "BUSY")
        
        # Extract organizer email and name
        organizer_email = self.extract_email_from_contact(organizer)
        organizer_name = self.extract_name_from_contact(organizer)
        
        # Add main event
        ics_content.extend([
            "BEGIN:VEVENT",
            f"UID:{uid}",
            "SEQUENCE:0",
            f"DTSTAMP:{datetime.now().strftime('%Y%m%dT%H%M%SZ')}",
            f"SUMMARY:{subject}",
            f"DESCRIPTION:{description.replace(chr(10), '\\n')}",
            f"LOCATION:{location}",
            f"ORGANIZER;CN={organizer_name}:mailto:{organizer_email}",
            f"PRIORITY:{ics_priority}",
            dtstart_line,
            dtend_line,
            "TRANSP:OPAQUE",
            f"X-MICROSOFT-CDO-BUSYSTATUS:{busy_status}",
            f"X-MICROSOFT-DISALLOW-COUNTER:{'TRUE' if not allow_counter else 'FALSE'}",
            f"LAST-MODIFIED:{datetime.now().strftime('%Y%m%dT%H%M%SZ')}"
        ])
        
        # Add attendees
        for attendee_contact in attendees:
            attendee_email = self.extract_email_from_contact(attendee_contact)
            attendee_name = self.extract_name_from_contact(attendee_contact)
            if attendee_email and '@' in attendee_email:
                ics_content.append(f"ATTENDEE;CN={attendee_name}:mailto:{attendee_email}")
        
        # Add privacy classification
        if private:
            ics_content.append("CLASS:PRIVATE")
        
        # Add reminder/alarm
        if reminder_enabled:
            reminder_map = {
                "15 minutes before": "-PT15M",
                "30 minutes before": "-PT30M", 
                "1 hour before": "-PT1H",
                "2 hours before": "-PT2H",
                "1 day before": "-P1D",
                "1 week before": "-P1W"
            }
            
            trigger = reminder_map.get(reminder_time, "-PT30M")
            ics_content.extend([
                "BEGIN:VALARM",
                "ACTION:DISPLAY",
                f"TRIGGER;RELATED=START:{trigger}",
                f"X-WR-ALARMUID:{str(uuid.uuid4())}",
                "END:VALARM"
            ])
        
        ics_content.extend([
            "END:VEVENT",
            "END:VCALENDAR"
        ])
        
        # Write to file
        with open(filename, 'w', encoding='utf-8') as f:
            f.write('\n'.join(ics_content))
    
    def toggle_time_controls(self, all_day, time_controls):
        """
        Enable/disable time controls based on all-day checkbox
        """
        state = "disabled" if all_day else "normal"
        for control in time_controls:
            control.config(state=state)
    
    # ========================================================================
    # TASK EDITING FUNCTIONALITY
    # ========================================================================
    def edit_selected_task(self):
        """
        Load selected task into the form for editing
        """
        selected_items = self.task_tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "Please select a task to edit.")
            return
        
        # Get the first selected task
        item = self.task_tree.item(selected_items[0])
        if not item["tags"]:  # Skip if it's a category, not a task
            messagebox.showwarning("Warning", "Please select a task, not a category.")
            return
        
        task_id = int(item["tags"][0])
        task = next((t for t in self.tasks if t["id"] == task_id), None)
        if not task:
            return
        
        # Load task data into the form
        self.populate_form_with_task(task)
        
        # Switch to edit mode
        self.editing_task = task
        self.create_update_button.config(text="Update Task")
        
        # Show the form if task list is visible
        if not self.task_list_visible:
            self.toggle_task_list()
        
        messagebox.showinfo("Edit Mode", f"Now editing task: '{task['subject']}'\nMake your changes and click 'Update Task'.")
    
    def populate_form_with_task(self, task):
        """
        Fill the form fields with data from an existing task
        """
        # Basic task information
        self.subject_var.set(task.get('subject', ''))
        self.priority_var.set(task.get('priority', 'None'))
        self.status_var.set(task.get('status', 'Not Started'))
        self.progress_var.set(task.get('progress', 0))
        self.update_progress_label(task.get('progress', 0))
        
        # Dates
        if task.get('start_date'):
            # Try to find matching start date option
            start_date_found = False
            for option, date_obj in self.start_options.items():
                if date_obj and option != "Custom":
                    if date_obj.strftime("%Y-%m-%d") == task['start_date']:
                        self.start_date_var.set(option)
                        start_date_found = True
                        break
            
            if not start_date_found:
                # Use custom date
                self.start_date_var.set('Custom')
                self.custom_start_date_var.set(self.convert_iso_to_display(task['start_date']))
                self.update_start_date()  # Show custom date field
        else:
            self.start_date_var.set('None')
        
        if task.get('due_date'):
            due_date_type = task.get('due_date_type', 'Custom')
            if due_date_type in self.due_options:
                self.due_date_var.set(due_date_type)
            else:
                self.due_date_var.set('Custom')
                self.custom_due_date_var.set(self.convert_iso_to_display(task['due_date']))
                self.update_due_date()  # Show custom date field
        else:
            self.due_date_var.set('None')
        
        # Reminders
        self.reminder_var.set(task.get('reminder', False))
        if task.get('reminder_date'):
            self.reminder_date_var.set(self.convert_iso_to_display(task['reminder_date']))
        
        if task.get('reminder_time'):
            time_parts = task['reminder_time'].split(':')
            if len(time_parts) >= 2:
                self.reminder_hour_var.set(time_parts[0])
                self.reminder_minute_var.set(time_parts[1])
        
        # Repeat settings
        repeat_settings = task.get('repeat_settings')
        if repeat_settings:
            self.repeat_frequency_var.set(repeat_settings.get('frequency', 'None'))
            self.repeat_end_var.set(repeat_settings.get('end_type', 'Never'))
            
            if repeat_settings.get('count'):
                self.repeat_count_var.set(str(repeat_settings['count']))
            if repeat_settings.get('end_date'):
                self.repeat_end_date_var.set(self.convert_iso_to_display(repeat_settings['end_date']))
        
        # Update reminder controls state
        self.toggle_reminder_options()
        self.update_repeat_options()
        self.update_repeat_end_options()
        
        # Organization
        project = task.get('project', '')
        if project and project in self.projects:
            self.project_var.set(project)
        else:
            self.project_var.set("None")
        
        category = task.get('category', '')
        if category and category in self.categories:
            self.category_var.set(category)
        else:
            self.category_var.set("-- Select Category --")
        
        # Handle owner field - support both old format (just name) and new format (Name <email>)
        owner_value = task.get('owner', '')
        if owner_value:
            # Check if it's already in "Name <email>" format
            if '<' not in owner_value and '@' not in owner_value:
                # It's just a name, try to find matching contact
                matches = self.find_contact_by_name_or_email(owner_value)
                if matches:
                    # Use the first match in full format
                    contact = matches[0]
                    owner_value = f"{contact['name']} <{contact['email']}>"
        self.owner_var.set(owner_value)
        
        self.location_var.set(task.get('location', ''))
        
        # Description
        self.description_text.delete("1.0", tk.END)
        if task.get('description'):
            self.description_text.insert(tk.END, task['description'])
    
    def create_or_update_task(self):
        """
        Create a new task or update an existing task based on current mode
        """
        if self.editing_task:
            self.update_task()
        else:
            self.create_task()
    
    def update_task(self):
        """
        Update the currently edited task with form data
        """
        # Validate required fields
        if not self.subject_var.get().strip():
            messagebox.showerror("Error", "Task name is required!")
            return
        
        # Validate category is selected
        category = self.category_var.get().strip()
        if not category or category == "-- Select Category --":
            messagebox.showerror("Error", "Please select a category for this task!")
            return
        
        # Handle new category/location/project
        if category and category not in self.categories:
            self.categories.append(category)
            self.categories.sort()
            self.category_combo.configure(values=["-- Select Category --"] + self.categories)
            self.save_categories()
        
        # Handle new project
        project = self.project_var.get().strip()
        if project and project != "None" and project not in self.projects:
            self.projects.append(project)
            self.projects.sort()
            self.project_combo.configure(values=["None"] + self.projects)
            self.save_projects()
        
        location = self.location_var.get().strip()
        if location and location not in self.locations:
            self.locations.append(location)
            self.locations.sort()
            self.location_combo.configure(values=self.locations)
            self.save_locations()
        
        # Update the task with form data
        self.editing_task.update({
            "subject": self.subject_var.get().strip(),
            "project": project if project != "None" else "",
            "priority": self.priority_var.get(),
            "status": self.status_var.get(),
            "progress": self.progress_var.get(),
            "start_date": self.get_actual_start_date(),
            "due_date": self.get_actual_due_date(),
            "due_date_type": self.due_date_var.get(),
            "reminder": self.reminder_var.get(),
            "reminder_date": self.convert_display_to_iso(self.reminder_date_var.get()) if self.reminder_var.get() else "",
            "reminder_time": self.get_reminder_time_string(),
            "repeat_settings": self.get_repeat_settings(),
            "category": category,
            "owner": self.owner_var.get().strip(),  # Store full contact format if available
            "location": location,
            "description": self.description_text.get("1.0", tk.END).strip(),
        })
        
        # Save and refresh
        self.auto_save_tasks()
        messagebox.showinfo("Success", f"Task '{self.editing_task['subject']}' updated successfully!")
        
        # Exit edit mode
        self.exit_edit_mode()
        self.refresh_task_list()
    
    def exit_edit_mode(self):
        """
        Exit edit mode and return to create mode
        """
        self.editing_task = None
        self.create_update_button.config(text="Create Task")
        self.clear_form()
    
    def toggle_task_list(self):
        """
        Toggle visibility of the task list panel and resize window accordingly
        """
        if self.task_list_visible:
            # Hide the task list
            self.tasks_frame.pack_forget()
            self.toggle_button.config(text="▶ Show Tasks")
            self.task_list_visible = False
            # Resize window to compact size
            self.root.geometry(self.compact_window_size)
            # Ensure form fills the available space
            self.main_frame.pack_configure(fill=tk.BOTH, expand=True)
        else:
            # Resize window back to full size first
            self.root.geometry(self.full_window_size)
            # Show the task list
            self.tasks_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0))
            self.toggle_button.config(text="◀ Hide Tasks")
            self.task_list_visible = True
    
    # ========================================================================
    # FORM INPUT HANDLING AND VALIDATION
    # ========================================================================
    def update_start_date(self, event=None):
        """
        Handle start date dropdown changes - show/hide custom date field
        """
        selected = self.start_date_var.get()
        if selected == "Custom":
            # Show custom date entry field
            self.custom_start_frame.grid(row=0, column=2, sticky=tk.W, pady=2, padx=(5, 0))
        else:
            # Hide custom field
            self.custom_start_frame.grid_remove()
    
    def update_due_date(self, event=None):
        """
        Handle due date dropdown changes - show/hide custom date field
        """
        selected = self.due_date_var.get()
        if selected == "Custom":
            # Show custom date entry field
            self.custom_due_frame.grid(row=1, column=2, sticky=tk.W, pady=2, padx=(5, 0))
        else:
            # Hide custom field and calculate date automatically
            self.custom_due_frame.grid_remove()
    
    def get_actual_start_date(self):
        """
        Calculate the actual start date based on dropdown selection or custom date
        """
        selected = self.start_date_var.get()
        if selected == "None":
            return ""
        elif selected == "Custom":
            return self.convert_display_to_iso(self.custom_start_date_var.get())
        elif selected in self.start_options:
            date_obj = self.start_options[selected]
            if date_obj:
                return date_obj.strftime("%Y-%m-%d")
        return ""
    
    def get_actual_due_date(self):
        """
        Calculate the actual due date based on dropdown selection or custom date
        """
        selected = self.due_date_var.get()
        if selected == "None":
            return ""
        elif selected == "Custom":
            return self.convert_display_to_iso(self.custom_due_date_var.get())
        elif selected in self.due_options:
            date_obj = self.due_options[selected]
            if date_obj:
                return date_obj.strftime("%Y-%m-%d")
        return ""
    
    def update_progress_label(self, value):
        """
        Update the progress percentage label when slider moves
        """
        self.progress_label.config(text=f"{int(float(value))}%")
    
    # ========================================================================
    # CALENDAR POPUP FUNCTIONALITY
    # ========================================================================
    def show_calendar(self, date_var, title):
        """
        Show a calendar popup for date selection
        """
        # Create calendar popup window
        cal_window = tk.Toplevel(self.root)
        cal_window.title(f"Select {title}")
        cal_window.resizable(False, False)
        cal_window.grab_set()  # Make window modal
        
        # Initialize calendar variables
        try:
            current_date = datetime.strptime(date_var.get(), "%b %d, %Y")
            self.cal_month = current_date.month
            self.cal_year = current_date.year
        except:
            today = datetime.now()
            self.cal_month = today.month
            self.cal_year = today.year
        
        # Header with navigation
        header_frame = ttk.Frame(cal_window)
        header_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(header_frame, text="<", width=3,
                  command=lambda: self.change_month(-1, cal_window, date_var)).pack(side=tk.LEFT)
        
        self.month_year_label = ttk.Label(header_frame, text="", font=("Arial", 12, "bold"))
        self.month_year_label.pack(side=tk.LEFT, expand=True)
        
        ttk.Button(header_frame, text=">", width=3,
                  command=lambda: self.change_month(1, cal_window, date_var)).pack(side=tk.RIGHT)
        
        # Calendar frame
        self.cal_frame = ttk.Frame(cal_window)
        self.cal_frame.pack(padx=10, pady=5)
        
        # Buttons frame
        buttons_frame = ttk.Frame(cal_window)
        buttons_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(buttons_frame, text="Today", 
                  command=lambda: self.select_today(cal_window, date_var)).pack(side=tk.LEFT)
        
        ttk.Button(buttons_frame, text="Cancel",
                  command=cal_window.destroy).pack(side=tk.RIGHT)
        
        # Build initial calendar
        self.build_calendar(cal_window, date_var)
        
        # Center the window
        cal_window.update_idletasks()
        x = (cal_window.winfo_screenwidth() // 2) - (cal_window.winfo_width() // 2)
        y = (cal_window.winfo_screenheight() // 2) - (cal_window.winfo_height() // 2)
        cal_window.geometry(f"+{x}+{y}")
    
    def change_month(self, direction, cal_window, date_var):
        """
        Navigate to previous/next month in calendar
        """
        self.cal_month += direction
        if self.cal_month > 12:
            self.cal_month = 1
            self.cal_year += 1
        elif self.cal_month < 1:
            self.cal_month = 12
            self.cal_year -= 1
        
        self.build_calendar(cal_window, date_var)
    
    def build_calendar(self, cal_window, date_var):
        """
        Build the calendar grid for the current month/year
        """
        # Clear existing calendar
        for widget in self.cal_frame.winfo_children():
            widget.destroy()
        
        # Update month/year label
        month_name = calendar.month_name[self.cal_month]
        self.month_year_label.config(text=f"{month_name} {self.cal_year}")
        
        # Day headers
        days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
        for i, day in enumerate(days):
            ttk.Label(self.cal_frame, text=day, font=("Arial", 9, "bold")).grid(row=0, column=i, padx=2, pady=2)
        
        # Get calendar data
        cal = calendar.monthcalendar(self.cal_year, self.cal_month)
        
        # Current date for highlighting
        today = datetime.now()
        
        # Build calendar grid
        for week_num, week in enumerate(cal, 1):
            for day_num, day in enumerate(week):
                if day == 0:
                    # Empty cells for days outside current month
                    ttk.Label(self.cal_frame, text="").grid(row=week_num, column=day_num, padx=2, pady=2)
                else:
                    # Regular day button
                    style = "normal"
                    if (self.cal_year == today.year and self.cal_month == today.month and day == today.day):
                        style = "bold"
                    
                    day_button = tk.Button(self.cal_frame, text=str(day), width=3, height=1,
                                         font=("Arial", 9, style),
                                         command=lambda d=day: self.select_date(d, cal_window, date_var))
                    day_button.grid(row=week_num, column=day_num, padx=1, pady=1)
    
    def select_date(self, day, cal_window, date_var):
        """
        Handle date selection from calendar
        """
        selected_date = datetime(self.cal_year, self.cal_month, day)
        date_var.set(selected_date.strftime("%b %d, %Y"))
        cal_window.destroy()
    
    def select_today(self, cal_window, date_var):
        """
        Select today's date and close calendar
        """
        today = datetime.now()
        date_var.set(today.strftime("%b %d, %Y"))
        cal_window.destroy()
    
    # ========================================================================
    # TASK MANAGEMENT FUNCTIONS
    # ========================================================================
    def create_task(self):
        """
        Create a new task from form data and save to storage
        """
        # Validate required fields
        if not self.subject_var.get().strip():
            messagebox.showerror("Error", "Task name is required!")
            return
        
        # Validate category is selected
        category = self.category_var.get().strip()
        if not category or category == "-- Select Category --":
            messagebox.showerror("Error", "Please select a category for this task!")
            return
        
        # Handle new category - add to list if it doesn't exist
        if category and category not in self.categories:
            self.categories.append(category)
            self.categories.sort()
            self.category_combo.configure(values=["-- Select Category --"] + self.categories)
            self.save_categories()
        
        # Handle new project - add to list if it doesn't exist
        project = self.project_var.get().strip()
        if project and project != "None" and project not in self.projects:
            self.projects.append(project)
            self.projects.sort()
            self.project_combo.configure(values=["None"] + self.projects)
            self.save_projects()
        
        # Handle new location - add to list if it doesn't exist
        location = self.location_var.get().strip()
        if location and location not in self.locations:
            self.locations.append(location)
            self.locations.sort()
            self.location_combo.configure(values=self.locations)
            self.save_locations()
        
        # Build task data dictionary from form inputs
        task = {
            "id": len(self.tasks) + 1,
            "subject": self.subject_var.get().strip(),
            "project": project if project != "None" else "",
            "priority": self.priority_var.get(),
            "status": self.status_var.get(),
            "progress": self.progress_var.get(),
            "start_date": self.get_actual_start_date(),
            "due_date": self.get_actual_due_date(),
            "due_date_type": self.due_date_var.get(),
            "reminder": self.reminder_var.get(),
            "reminder_date": self.convert_display_to_iso(self.reminder_date_var.get()) if self.reminder_var.get() else "",
            "reminder_time": self.get_reminder_time_string(),
            "repeat_settings": self.get_repeat_settings(),
            "category": category,
            "owner": self.owner_var.get().strip(),  # Store full contact format if available
            "location": location,
            "description": self.description_text.get("1.0", tk.END).strip(),
            "created_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        # Add task to list and auto-save
        self.tasks.append(task)
        self.auto_save_tasks()  # Automatically save after creating task
        
        # Show task list if it's hidden so user can see the new task
        if not self.task_list_visible:
            self.toggle_task_list()
        
        messagebox.showinfo("Success", f"Task '{task['subject']}' created and saved successfully!")
        
        # Clear form for next task
        self.clear_form()
        
        # Refresh task list to show new task
        self.refresh_task_list()
    
    def clear_form(self):
        """
        Reset all form fields to default values and exit edit mode
        """
        # Exit edit mode if active
        if self.editing_task:
            self.exit_edit_mode()
            return
        
        # Clear basic task fields
        self.subject_var.set("")
        self.priority_var.set("None")
        self.status_var.set("Not Started")
        self.progress_var.set(0)
        self.update_progress_label(0)
        
        # Reset dates to default values
        self.start_date_var.set("None")
        self.due_date_var.set("None")
        today = datetime.now().strftime("%b %d, %Y")
        self.custom_start_date_var.set(today)
        self.custom_due_date_var.set(today)
        self.custom_start_frame.grid_remove()  # Hide custom start date
        self.custom_due_frame.grid_remove()  # Hide custom due date
        
        # Clear reminder settings
        self.reminder_var.set(False)
        self.reminder_date_var.set(today)
        self.reminder_hour_var.set("09")
        self.reminder_minute_var.set("00")
        self.repeat_frequency_var.set("None")
        self.repeat_end_var.set("Never")
        self.repeat_count_var.set("5")
        self.repeat_end_date_var.set(today)
        
        # Reset reminder controls state
        self.toggle_reminder_options()
        
        # Clear organization fields
        self.project_var.set("None")
        self.category_var.set("-- Select Category --")
        self.owner_var.set("")
        self.location_var.set("")
        
        # Clear description
        self.description_text.delete("1.0", tk.END)
    
    def refresh_task_list(self):
        """
        Reload and display all tasks in the tree view with category grouping
        """
        # Clear existing tree items
        for item in self.task_tree.get_children():
            self.task_tree.delete(item)
        
        # Group tasks by category
        categories = {}
        for task in self.tasks:
            category = task.get('category', 'Unknown')  # Should not happen with required categories
            if category not in categories:
                categories[category] = []
            categories[category].append(task)
        
        # Add categories and tasks to tree
        for category, tasks in sorted(categories.items()):
            # Add category header
            category_item = self.task_tree.insert("", "end", text=f"📁 {category}", open=True)
            
            # Add tasks under category
            for task in sorted(tasks, key=lambda x: x.get('due_date', '')):
                # Format due date for display
                due_date = task.get('due_date', '')
                if due_date:
                    try:
                        due_obj = datetime.strptime(due_date, "%Y-%m-%d")
                        due_display = due_obj.strftime("%b %d")
                    except:
                        due_display = due_date
                else:
                    due_display = ""
                
                # Format owner info for display
                owner_info = task.get('owner', '')
                if owner_info:
                    # Extract just the name from "Name <email>" format
                    if '<' in owner_info and '>' in owner_info:
                        owner_display = owner_info.split('<')[0].strip()
                    else:
                        owner_display = owner_info
                else:
                    owner_display = ""
                
                # Add task to tree with updated columns and project formatting
                task_display_name = self.format_task_name_with_project(task['subject'], task.get('project', ''))
                self.task_tree.insert(category_item, "end", 
                                    text=f"📋 {task_display_name}", 
                                    values=(task['priority'], task['status'], due_display, owner_display),
                                    tags=(str(task['id']),))
    
    def delete_selected_task(self):
        """
        Delete selected task(s) from the list with confirmation
        """
        selected_items = self.task_tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "Please select a task to delete.")
            return
        
        # Filter out categories and get only tasks
        selected_tasks = []
        for item_id in selected_items:
            item = self.task_tree.item(item_id)
            # Only include items that have tags (tasks, not categories)
            if item["tags"]:
                task_id = int(item["tags"][0])
                task = next((t for t in self.tasks if t["id"] == task_id), None)
                if task:
                    selected_tasks.append(task)
        
        if not selected_tasks:
            messagebox.showwarning("Warning", "Please select actual tasks, not categories.")
            return
        
        # Confirmation dialog
        task_names = [task['subject'] for task in selected_tasks]
        if len(task_names) == 1:
            message = f"Are you sure you want to delete task '{task_names[0]}'?"
        else:
            message = f"Are you sure you want to delete {len(task_names)} selected tasks?"
        
        if messagebox.askyesno("Confirm Deletion", message):
            # Remove tasks from list
            for task in selected_tasks:
                self.tasks.remove(task)
            
            # Auto-save and refresh
            self.auto_save_tasks()
            self.refresh_task_list()
            messagebox.showinfo("Success", f"{len(selected_tasks)} task(s) deleted successfully!")
    
    def view_task_details(self, event):
        """
        Show detailed task information in a popup window - no scrollbars, shorter window
        """
        selected_item = self.task_tree.selection()
        if not selected_item:
            return
        
        item = self.task_tree.item(selected_item[0])
        if not item["tags"]:  # Skip if it's a category, not a task
            return
        
        task_id = int(item["tags"][0])
        task = next((t for t in self.tasks if t["id"] == task_id), None)
        if not task:
            return
        
        # Create details popup window - smaller and no scrollbars
        details_window = tk.Toplevel(self.root)
        details_window.title(f"Task Details: {task['subject']}")
        details_window.geometry("500x450")  # Reduced height from 600 to 450
        details_window.resizable(False, False)
        
        # Simple frame without scrolling - reduced padding
        details_frame = ttk.Frame(details_window)
        details_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)  # Reduced from 20 to 10
        
        # Basic information
        info_items = [
            ("Task", self.format_task_name_with_project(task['subject'], task.get('project', ''))),
            ("Project", task.get('project', 'None') if task.get('project') else 'None'),
            ("Priority", task['priority']),
            ("Status", task['status']),
            ("Progress", f"{task['progress']}%"),
            ("Start Date", self.convert_iso_to_display(task['start_date'])),
            ("Due Date", self.convert_iso_to_display(task['due_date'])),
            ("Category", task.get('category', '')),
            ("Owner", task.get('owner', '')),
            ("Location", task.get('location', '')),
        ]
        
        for i, (label, value) in enumerate(info_items):
            ttk.Label(details_frame, text=f"{label}:", font=("Arial", 10, "bold")).grid(
                row=i, column=0, sticky=tk.W, pady=1, padx=(0, 8))  # Reduced from pady=2, padx=(0, 10)
            ttk.Label(details_frame, text=str(value), font=("Arial", 10)).grid(
                row=i, column=1, sticky=tk.W, pady=1)  # Reduced from pady=2
        
        # Reminder details
        row_offset = len(info_items)
        if task.get('reminder'):
            ttk.Label(details_frame, text="Reminder:", font=("Arial", 10, "bold")).grid(
                row=row_offset, column=0, sticky=tk.W, pady=1, padx=(0, 8))  # Reduced padding
            
            reminder_text = "Yes"
            if task.get('reminder_date'):
                reminder_text += f" on {self.convert_iso_to_display(task['reminder_date'])}"
            if task.get('reminder_time'):
                reminder_text += f" at {task['reminder_time']}"
            
            ttk.Label(details_frame, text=reminder_text, font=("Arial", 10)).grid(
                row=row_offset, column=1, sticky=tk.W, pady=1)  # Reduced padding
            row_offset += 1
            
            # Repeat settings
            if task.get('repeat_settings'):
                repeat_settings = task['repeat_settings']
                ttk.Label(details_frame, text="Repeat:", font=("Arial", 10, "bold")).grid(
                    row=row_offset, column=0, sticky=tk.W, pady=1, padx=(0, 8))  # Reduced padding
                
                repeat_text = repeat_settings.get('frequency', 'None')
                if repeat_settings.get('end_type') == "After X occurrences":
                    repeat_text += f" (ends after {repeat_settings.get('count', 0)} times)"
                elif repeat_settings.get('end_type') == "Until date":
                    end_date = repeat_settings.get('end_date', '')
                    if end_date:
                        repeat_text += f" (until {self.convert_iso_to_display(end_date)})"
                
                ttk.Label(details_frame, text=repeat_text, font=("Arial", 10)).grid(
                    row=row_offset, column=1, sticky=tk.W, pady=1)  # Reduced padding
                row_offset += 1
        
        # Description - smaller text area
        if task.get('description'):
            ttk.Label(details_frame, text="Description:", font=("Arial", 10, "bold")).grid(
                row=row_offset, column=0, sticky=tk.NW, pady=(8, 1), padx=(0, 8))  # Reduced padding
            
            desc_text = tk.Text(details_frame, height=5, width=50, wrap=tk.WORD, 
                              font=("Arial", 10), state=tk.NORMAL)
            desc_text.grid(row=row_offset, column=1, sticky=(tk.W, tk.E), pady=(8, 1))  # Reduced padding
            desc_text.insert(tk.END, task['description'])
            desc_text.config(state=tk.DISABLED)
        
        # Close button
        ttk.Button(details_frame, text="Close", command=details_window.destroy).grid(
            row=row_offset + 1, column=0, columnspan=2, pady=15)  # Reduced from pady=20
        
        # Center the window
        details_window.update_idletasks()
        x = (details_window.winfo_screenwidth() // 2) - (details_window.winfo_width() // 2)
        y = (details_window.winfo_screenheight() // 2) - (details_window.winfo_height() // 2)
        details_window.geometry(f"+{x}+{y}")
    
    # ========================================================================
    # FILE OPERATIONS AND DATA PERSISTENCE
    # ========================================================================
    def load_tasks(self):
        """
        Load tasks from JSON file if it exists
        """
        try:
            if os.path.exists("tasks.json"):
                with open("tasks.json", "r") as f:
                    self.tasks = json.load(f)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load tasks: {e}")
            self.tasks = []
    
    def auto_save_tasks(self):
        """
        Automatically save tasks to JSON file
        """
        try:
            with open("tasks.json", "w") as f:
                json.dump(self.tasks, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save tasks: {e}")
    
    def load_categories(self):
        """
        Load categories from JSON file if it exists
        """
        try:
            if os.path.exists("categories.json"):
                with open("categories.json", "r") as f:
                    loaded_categories = json.load(f)
                    # Merge with defaults, removing duplicates
                    all_categories = list(set(self.categories + loaded_categories))
                    self.categories = sorted(all_categories)
        except Exception as e:
            pass  # Use defaults if loading fails
    
    def save_categories(self):
        """
        Save categories to JSON file
        """
        try:
            with open("categories.json", "w") as f:
                json.dump(self.categories, f, indent=2)
        except Exception as e:
            pass  # Fail silently for categories
    
    def load_locations(self):
        """
        Load locations from JSON file if it exists
        """
        try:
            if os.path.exists("locations.json"):
                with open("locations.json", "r") as f:
                    self.locations = json.load(f)
        except Exception as e:
            pass  # Use empty list if loading fails
    
    def save_locations(self):
        """
        Save locations to JSON file
        """
        try:
            with open("locations.json", "w") as f:
                json.dump(self.locations, f, indent=2)
        except Exception as e:
            pass  # Fail silently for locations
    
    def load_contacts(self):
        """
        Load contacts from JSON file if it exists
        """
        try:
            if os.path.exists("contacts.json"):
                with open("contacts.json", "r") as f:
                    self.contacts = json.load(f)
            else:
                # Default contacts - add yourself as first contact
                self.contacts = [
                    {"name": "Randy Buhr", "email": "rabuhr@fseinc.net"},
                ]
        except Exception as e:
            self.contacts = [{"name": "Randy Buhr", "email": "rabuhr@fseinc.net"}]
    
    def save_contacts(self):
        """
        Save contacts to JSON file
        """
        try:
            with open("contacts.json", "w") as f:
                json.dump(self.contacts, f, indent=2)
        except Exception as e:
            pass  # Fail silently for contacts
    
    def load_projects(self):
        """
        Load projects from JSON file if it exists
        """
        try:
            if os.path.exists("projects.json"):
                with open("projects.json", "r") as f:
                    self.projects = json.load(f)
        except Exception as e:
            pass  # Use empty list if loading fails
    
    def save_projects(self):
        """
        Save projects to JSON file
        """
        try:
            with open("projects.json", "w") as f:
                json.dump(self.projects, f, indent=2)
        except Exception as e:
            pass  # Fail silently for projects
    
    def show_project_manager(self, parent_window):
        """
        Show project management dialog
        """
        # Create project manager window
        project_window = tk.Toplevel(parent_window)
        project_window.title("Manage Projects")
        project_window.geometry("500x400")
        project_window.resizable(False, False)
        project_window.grab_set()  # Make window modal
        
        # Main frame with reduced padding
        main_frame = ttk.Frame(project_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        
        # Header
        ttk.Label(main_frame, text="Project Manager", font=("Arial", 12, "bold")).pack(anchor=tk.W, pady=(0, 5))
        
        # Project list
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        # Listbox with scrollbar
        project_listbox = tk.Listbox(list_frame, height=15)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=project_listbox.yview)
        project_listbox.configure(yscrollcommand=scrollbar.set)
        
        project_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Populate project list
        def refresh_project_list():
            project_listbox.delete(0, tk.END)
            for project in self.projects:
                project_listbox.insert(tk.END, project)
        
        refresh_project_list()
        
        # Add new project frame
        add_frame = ttk.LabelFrame(main_frame, text="Add New Project", padding="5")
        add_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(add_frame, text="Project Name:").grid(row=0, column=0, sticky=tk.W, pady=2)
        new_project_var = tk.StringVar()
        project_entry = ttk.Entry(add_frame, textvariable=new_project_var, width=30)
        project_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        
        def add_project():
            project_name = new_project_var.get().strip()
            if project_name:
                # Check for duplicates
                if project_name in self.projects:
                    messagebox.showwarning("Duplicate", f"Project '{project_name}' already exists!")
                    return
                
                self.projects.append(project_name)
                self.projects.sort()
                self.save_projects()
                refresh_project_list()
                
                # Update the main form's project combo
                self.project_combo.configure(values=["None"] + self.projects)
                
                new_project_var.set("")
                messagebox.showinfo("Success", f"Added project '{project_name}'!")
            else:
                messagebox.showwarning("Invalid Input", "Please enter a project name.")
        
        ttk.Button(add_frame, text="Add Project", command=add_project).grid(row=0, column=2, padx=(5, 0))
        add_frame.columnconfigure(1, weight=1)
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X)
        
        def delete_selected():
            selection = project_listbox.curselection()
            if selection:
                index = selection[0]
                project_name = self.projects[index]
                if messagebox.askyesno("Confirm Delete", f"Delete project '{project_name}'?\n\nNote: Tasks using this project will show '[{project_name}]' in their names until you edit them."):
                    del self.projects[index]
                    self.save_projects()
                    refresh_project_list()
                    
                    # Update the main form's project combo
                    self.project_combo.configure(values=["None"] + self.projects)
            else:
                messagebox.showwarning("No Selection", "Please select a project to delete.")
        
        ttk.Button(buttons_frame, text="Delete Selected", command=delete_selected).pack(side=tk.LEFT, padx=3)
        ttk.Button(buttons_frame, text="Close", command=project_window.destroy).pack(side=tk.RIGHT, padx=3)
        
        # Bind Enter key to add project
        project_entry.bind('<Return>', lambda e: add_project())
        
        # Center the window
        project_window.update_idletasks()
        x = (project_window.winfo_screenwidth() // 2) - (project_window.winfo_width() // 2)
        y = (project_window.winfo_screenheight() // 2) - (project_window.winfo_height() // 2)
        project_window.geometry(f"+{x}+{y}")
    
    def find_contact_by_name_or_email(self, search_text):
        """
        Find contacts that match the search text (name or email)
        """
        search_text = search_text.lower()
        matches = []
        for contact in self.contacts:
            if (search_text in contact['name'].lower() or 
                search_text in contact['email'].lower()):
                matches.append(contact)
        return matches
    
    def format_task_name_with_project(self, task_subject, project):
        """
        Format task name with project prefix for display and export
        """
        if project and project != "None":
            return f"[{project}] {task_subject}"
        return task_subject
        """
        Find contacts that match the search text (name or email)
        """
        search_text = search_text.lower()
        matches = []
        for contact in self.contacts:
            if (search_text in contact['name'].lower() or 
                search_text in contact['email'].lower()):
                matches.append(contact)
        return matches
    
    def show_contact_manager(self, parent_window):
        """
        Show contact management dialog - reduced padding
        """
        # Create contact manager window
        contact_window = tk.Toplevel(parent_window)
        contact_window.title("Manage Contacts")
        contact_window.geometry("500x500")
        contact_window.resizable(False, False)
        contact_window.grab_set()  # Make window modal
        
        # Main frame with reduced padding
        main_frame = ttk.Frame(contact_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)  # Reduced from 15 to 8
        
        # Header
        ttk.Label(main_frame, text="Contact Manager", font=("Arial", 12, "bold")).pack(anchor=tk.W, pady=(0, 5))  # Reduced from 8 to 5
        
        # Contact list
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))  # Reduced from 8 to 5
        
        # Listbox with scrollbar
        contact_listbox = tk.Listbox(list_frame, height=18)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=contact_listbox.yview)
        contact_listbox.configure(yscrollcommand=scrollbar.set)
        
        contact_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Populate contact list
        def refresh_contact_list():
            contact_listbox.delete(0, tk.END)
            for contact in self.contacts:
                contact_listbox.insert(tk.END, f"{contact['name']} <{contact['email']}>")
        
        refresh_contact_list()
        
        # Add new contact frame with reduced padding
        add_frame = ttk.LabelFrame(main_frame, text="Add New Contact", padding="5")  # Reduced from 8 to 5
        add_frame.pack(fill=tk.X, pady=(0, 5))  # Reduced from 8 to 5
        
        ttk.Label(add_frame, text="Name:").grid(row=0, column=0, sticky=tk.W, pady=2)
        new_name_var = tk.StringVar()
        name_entry = ttk.Entry(add_frame, textvariable=new_name_var, width=25)
        name_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        
        ttk.Label(add_frame, text="Email:").grid(row=1, column=0, sticky=tk.W, pady=2)
        new_email_var = tk.StringVar()
        email_entry = ttk.Entry(add_frame, textvariable=new_email_var, width=25)
        email_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=2, padx=(5, 0))
        
        def add_contact():
            name = new_name_var.get().strip()
            email = new_email_var.get().strip()
            if name and email and '@' in email:
                # Check for duplicates
                for contact in self.contacts:
                    if contact['email'].lower() == email.lower():
                        messagebox.showwarning("Duplicate", f"Contact with email {email} already exists!")
                        return
                
                self.contacts.append({"name": name, "email": email})
                self.save_contacts()
                refresh_contact_list()
                new_name_var.set("")
                new_email_var.set("")
                messagebox.showinfo("Success", f"Added {name} to contacts!")
            else:
                messagebox.showwarning("Invalid Input", "Please enter both name and valid email address.")
        
        ttk.Button(add_frame, text="Add Contact", command=add_contact).grid(row=0, column=2, rowspan=2, padx=(5, 0))  # Reduced from 8 to 5
        add_frame.columnconfigure(1, weight=1)
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X)
        
        def delete_selected():
            selection = contact_listbox.curselection()
            if selection:
                index = selection[0]
                contact = self.contacts[index]
                if messagebox.askyesno("Confirm Delete", f"Delete contact {contact['name']}?"):
                    del self.contacts[index]
                    self.save_contacts()
                    refresh_contact_list()
            else:
                messagebox.showwarning("No Selection", "Please select a contact to delete.")
        
        ttk.Button(buttons_frame, text="Delete Selected", command=delete_selected).pack(side=tk.LEFT, padx=3)
        ttk.Button(buttons_frame, text="Close", command=contact_window.destroy).pack(side=tk.RIGHT, padx=3)
        
        # Center the window
        contact_window.update_idletasks()
        x = (contact_window.winfo_screenwidth() // 2) - (contact_window.winfo_width() // 2)
        y = (contact_window.winfo_screenheight() // 2) - (contact_window.winfo_height() // 2)
        contact_window.geometry(f"+{x}+{y}")
    
    def create_contact_autocomplete_entry(self, parent, textvariable, width=35):
        """
        Create an entry field with contact autocomplete functionality
        """
        frame = ttk.Frame(parent)
        
        # Entry field
        entry = ttk.Entry(frame, textvariable=textvariable, width=width)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Contact picker button
        def pick_contact():
            self.show_contact_picker(textvariable)
        
        pick_button = ttk.Button(frame, text="📋", width=3, command=pick_contact)
        pick_button.pack(side=tk.LEFT, padx=(1, 0))
        
        # Bind autocomplete
        def on_key_release(event):
            if event.keysym in ['Up', 'Down', 'Left', 'Right', 'Tab']:
                return
            
            current_text = textvariable.get()
            if len(current_text) < 2:
                return
            
            matches = self.find_contact_by_name_or_email(current_text)
            if matches:
                # Show first match as suggestion
                first_match = matches[0]
                suggestion = f"{first_match['name']} <{first_match['email']}>"
                
                # Simple autocomplete - you can enhance this with a dropdown later
                if current_text.lower() in first_match['name'].lower():
                    # User is typing a name, suggest the full contact
                    pass  # Could implement dropdown here
        
        entry.bind('<KeyRelease>', on_key_release)
        
        return frame, entry
    
    def show_contact_picker(self, target_var):
        """
        Show a contact picker dialog
        """
        # Create picker window
        picker = tk.Toplevel(self.root)
        picker.title("Select Contact")
        picker.geometry("400x300")
        picker.grab_set()
        
        # Contact list
        main_frame = ttk.Frame(picker)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        ttk.Label(main_frame, text="Select a contact:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))
        
        contact_listbox = tk.Listbox(main_frame, height=12)
        contact_listbox.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        for contact in self.contacts:
            contact_listbox.insert(tk.END, f"{contact['name']} <{contact['email']}>")
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        def select_contact():
            selection = contact_listbox.curselection()
            if selection:
                index = selection[0]
                contact = self.contacts[index]
                target_var.set(f"{contact['name']} <{contact['email']}>")
                picker.destroy()
        
        def add_new_contact():
            picker.destroy()
            self.show_contact_manager(self.root)
        
        ttk.Button(button_frame, text="Select", command=select_contact).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Add New", command=add_new_contact).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=picker.destroy).pack(side=tk.RIGHT, padx=5)
        
        # Double-click to select
        contact_listbox.bind('<Double-Button-1>', lambda e: select_contact())
        
        # Center window
        picker.update_idletasks()
        x = (picker.winfo_screenwidth() // 2) - (picker.winfo_width() // 2)
        y = (picker.winfo_screenheight() // 2) - (picker.winfo_height() // 2)
        picker.geometry(f"+{x}+{y}")
    
    def get_attendees_list(self, listbox):
        """
        Get all attendees from the listbox as a list
        """
        attendees = []
        for i in range(listbox.size()):
            attendees.append(listbox.get(i))
        return attendees
    
    def extract_email_from_contact(self, contact_string):
        """
        Extract email from "Name <email>" format or return as-is if just email
        """
        import re
        
        # Match "Name <email@domain.com>" format
        match = re.search(r'<([^>]+)>', contact_string)
        if match:
            return match.group(1).strip()
        
        # If no angle brackets, assume it's just an email
        if '@' in contact_string:
            return contact_string.strip()
        
        return contact_string.strip()
    
    def extract_name_from_contact(self, contact_string):
        """
        Extract name from "Name <email>" format or return email if no name
        """
        import re
        
        # Match "Name <email@domain.com>" format
        if '<' in contact_string and '>' in contact_string:
            name_part = contact_string.split('<')[0].strip()
            if name_part:
                return name_part
            # If no name part, extract name from email
            email = self.extract_email_from_contact(contact_string)
            return email.split('@')[0].replace('.', ' ').title()
        
        # If no angle brackets, extract name from email
        if '@' in contact_string:
            return contact_string.split('@')[0].replace('.', ' ').title()
        
        return contact_string.strip()
    
    # ========================================================================
    # EXPORT FUNCTIONALITY
    # ========================================================================
    def export_selected_task(self):
        """
        Export selected task(s) as iCalendar (.ics) file
        """
        selected_items = self.task_tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "Please select task(s) to export.")
            return
        
        # Filter out categories and get only tasks
        selected_tasks = []
        for item_id in selected_items:
            item = self.task_tree.item(item_id)
            # Only include items that have tags (tasks, not categories)
            if item["tags"]:
                task_id = int(item["tags"][0])
                task = next((t for t in self.tasks if t["id"] == task_id), None)
                if task:
                    selected_tasks.append(task)
        
        if not selected_tasks:
            messagebox.showwarning("Warning", "Please select actual tasks, not categories.")
            return
        
        # Create filename based on number of tasks
        if len(selected_tasks) == 1:
            task_name = self.format_task_name_with_project(selected_tasks[0]['subject'], selected_tasks[0].get('project', ''))
            safe_filename = "".join(c for c in task_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            default_name = f"{safe_filename}.ics"
        else:
            default_name = f"Tasks_Export_{len(selected_tasks)}_items.ics"
        
        # Show file save dialog for .ics file
        filename = filedialog.asksaveasfilename(
            defaultextension=".ics",
            initialfile=default_name,
            filetypes=[("iCalendar files", "*.ics"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                self.generate_ics_file(selected_tasks, filename)
                messagebox.showinfo("Success", f"Tasks exported successfully to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export tasks: {e}")
    
    def generate_ics_file(self, tasks, filename):
        """
        Generate iCalendar (.ics) file with enhanced reminder support
        """
        ics_content = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Exchange Task Creator//Enhanced//EN"]
        
        for task in tasks:
            # Create unique UID for each task
            uid = str(uuid.uuid4())
            
            # Basic task information with project formatting
            task_display_name = self.format_task_name_with_project(task['subject'], task.get('project', ''))
            ics_content.extend([
                "BEGIN:VTODO",
                f"UID:{uid}",
                f"SUMMARY:{task_display_name}",
                f"PRIORITY:{self.get_ics_priority(task['priority'])}",
                f"STATUS:{self.get_ics_status(task['status'])}",
                f"PERCENT-COMPLETE:{task['progress']}",
                f"DTSTAMP:{datetime.now().strftime('%Y%m%dT%H%M%SZ')}"
            ])
            
            # Dates
            if task.get('start_date'):
                start_date = task['start_date'].replace('-', '')
                ics_content.append(f"DTSTART;VALUE=DATE:{start_date}")
            
            if task.get('due_date'):
                due_date = task['due_date'].replace('-', '')
                ics_content.append(f"DUE;VALUE=DATE:{due_date}")
            
            # Enhanced reminder with repeat support
            if task.get('reminder') and task.get('reminder_date'):
                try:
                    reminder_datetime = datetime.strptime(task['reminder_date'], "%Y-%m-%d")
                    
                    # Add time if specified
                    if task.get('reminder_time'):
                        time_parts = task['reminder_time'].split(':')
                        reminder_datetime = reminder_datetime.replace(
                            hour=int(time_parts[0]), 
                            minute=int(time_parts[1])
                        )
                    
                    reminder_str = reminder_datetime.strftime('%Y%m%dT%H%M%SZ')
                    
                    # Add reminder alarm
                    ics_content.extend([
                        "BEGIN:VALARM",
                        "ACTION:DISPLAY",
                        f"DESCRIPTION:Reminder: {task['subject']}",
                        f"TRIGGER;VALUE=DATE-TIME:{reminder_str}"
                    ])
                    
                    # Add repeat rule if specified
                    if task.get('repeat_settings'):
                        repeat_rule = self.generate_rrule(task['repeat_settings'])
                        if repeat_rule:
                            ics_content.append(f"RRULE:{repeat_rule}")
                    
                    ics_content.append("END:VALARM")
                    
                except Exception as e:
                    pass  # Skip invalid reminder dates
            
            # Additional properties
            if task.get('category'):
                ics_content.append(f"CATEGORIES:{task['category']}")
            
            if task.get('location'):
                ics_content.append(f"LOCATION:{task['location']}")
            
            if task.get('description'):
                # Escape description for iCalendar format
                escaped_desc = task['description'].replace('\n', '\\n').replace(',', '\\,').replace(';', '\\;')
                ics_content.append(f"DESCRIPTION:{escaped_desc}")
            
            ics_content.append("END:VTODO")
        
        ics_content.append("END:VCALENDAR")
        
        # Write to file
        with open(filename, 'w') as f:
            f.write('\n'.join(ics_content))
    
    def generate_rrule(self, repeat_settings):
        """
        Generate RRULE string for repeating reminders
        """
        if not repeat_settings or repeat_settings.get('frequency') == 'None':
            return None
        
        frequency_map = {
            'Daily': 'DAILY',
            'Weekly': 'WEEKLY', 
            'Monthly': 'MONTHLY',
            'Yearly': 'YEARLY'
        }
        
        freq = frequency_map.get(repeat_settings['frequency'])
        if not freq:
            return None
        
        rrule_parts = [f"FREQ={freq}"]
        
        # Add end conditions
        end_type = repeat_settings.get('end_type', 'Never')
        if end_type == "After X occurrences" and repeat_settings.get('count'):
            rrule_parts.append(f"COUNT={repeat_settings['count']}")
        elif end_type == "Until date" and repeat_settings.get('end_date'):
            try:
                end_date = datetime.strptime(repeat_settings['end_date'], "%Y-%m-%d")
                rrule_parts.append(f"UNTIL={end_date.strftime('%Y%m%dT235959Z')}")
            except:
                pass
        
        return ';'.join(rrule_parts)
    
    def get_ics_priority(self, priority):
        """
        Convert priority to iCalendar format (1-9 scale)
        """
        priority_map = {"High": "1", "Medium": "5", "Low": "9", "None": "0"}
        return priority_map.get(priority, "0")
    
    def get_ics_status(self, status):
        """
        Convert status to iCalendar format
        """
        status_map = {
            "Not Started": "NEEDS-ACTION",
            "In Progress": "IN-PROCESS", 
            "Completed": "COMPLETED",
            "Waiting": "NEEDS-ACTION",
            "Deferred": "CANCELLED"
        }
        return status_map.get(status, "NEEDS-ACTION")
    
    # ========================================================================
    # UTILITY FUNCTIONS
    # ========================================================================
    def convert_display_to_iso(self, display_date):
        """
        Convert display format date (e.g., "Jan 15, 2024") to ISO format (e.g., "2024-01-15")
        """
        try:
            date_obj = datetime.strptime(display_date, "%b %d, %Y")
            return date_obj.strftime("%Y-%m-%d")
        except:
            return ""
    
    def convert_iso_to_display(self, iso_date):
        """
        Convert ISO format date (e.g., "2024-01-15") to display format (e.g., "Jan 15, 2024")
        """
        if not iso_date:
            return "None"
        try:
            date_obj = datetime.strptime(iso_date, "%Y-%m-%d")
            return date_obj.strftime("%b %d, %Y")
        except:
            return iso_date if iso_date else "None"

# ============================================================================
# APPLICATION ENTRY POINT
# ============================================================================
def main():
    """
    Create and run the main application
    """
    # Use TkinterDnD if available for drag and drop support
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    
    app = ExchangeTaskCreator(root)
    root.mainloop()

if __name__ == "__main__":
    main()
