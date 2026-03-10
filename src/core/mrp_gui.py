"""
MRP Critical Items Analyzer - GUI Module
A professional GUI application for analyzing MRP data and identifying critical items.

Author: Humberto Rodrigues
Date: August 2025
Version: 1.0.0
"""

import os
import time
import json
import webbrowser
from pathlib import Path
from datetime import datetime
from dataclasses import dataclass, field
from typing import Optional, Dict, Any, Tuple, List

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from ttkbootstrap import Style
from ttkbootstrap.tooltip import ToolTip

from src.core.mrp_analyzer import MRPAnalyzer, MRPConfig

# Configure logging
import logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)


@dataclass
class GUIConfig:
    """Configuration management for the GUI application."""
    # File and directory settings
    last_directory: str = field(default="")
    default_sheet_name: str = field(default="Cálculo MRP")
    
    # UI settings
    theme: str = field(default="flatly")
    window_size: Tuple[int, int] = field(default=(1200, 800))
    min_window_size: Tuple[int, int] = field(default=(900, 600))
    
    # Table settings
    page_size: int = field(default=50)
    table_columns: List[str] = field(default_factory=lambda: [
        "CÓD", "FORNECEDOR PRINCIPAL", "DESCRIÇÃOPROMOB", 
        "ESTOQUE DISPONÍVEL", "QUANTIDADE A SOLICITAR"
    ])
    
    # File paths
    config_dir: Path = field(default_factory=lambda: Path.home() / '.mrp_analyzer')
    config_file: Path = field(default_factory=lambda: Path.home() / '.mrp_analyzer' / 'config.json')
    
    def __post_init__(self):
        """Ensures configuration directory exists after initialization."""
        self.config_dir.mkdir(parents=True, exist_ok=True)
    
    @classmethod
    def load(cls) -> 'GUIConfig':
        """
        Loads configuration from file or creates default configuration.
        
        Returns:
            GUIConfig: Loaded or default configuration
        """
        default_config = cls()
        try:
            if default_config.config_file.exists():
                with open(default_config.config_file) as f:
                    raw_config = json.load(f)

                # Restore serialized tuple fields and keep a safe fallback for unknown keys
                if 'window_size' in raw_config:
                    raw_config['window_size'] = tuple(raw_config['window_size'])
                if 'min_window_size' in raw_config:
                    raw_config['min_window_size'] = tuple(raw_config['min_window_size'])
                if 'config_dir' in raw_config:
                    raw_config['config_dir'] = Path(raw_config['config_dir'])
                if 'config_file' in raw_config:
                    raw_config['config_file'] = Path(raw_config['config_file'])

                return cls(**raw_config)
        except Exception as e:
            logger.error(f"Error loading config: {e}")
        return default_config
    
    def save(self) -> None:
        """Saves current configuration to file."""
        try:
            payload = {
                **self.__dict__,
                'config_dir': str(self.config_dir),
                'config_file': str(self.config_file),
            }
            with open(self.config_file, 'w') as f:
                json.dump(payload, f, indent=2)
            logger.info("Configuration saved successfully")
        except Exception as e:
            logger.error(f"Error saving config: {e}")

@dataclass
class AppState:
    """Application state management."""
    # Configuration
    config: GUIConfig = field(default_factory=GUIConfig.load)
    
    # Data state
    df_table: pd.DataFrame = field(default_factory=pd.DataFrame)
    current_page: int = field(default=0)
    total_pages: int = field(default=0)
    
    # UI state
    filter_applied: bool = field(default=False)
    last_sort_column: Optional[str] = field(default=None)
    sort_ascending: bool = field(default=True)
    
    # Analysis state
    mrp_analyzer: MRPAnalyzer = field(default_factory=MRPAnalyzer)
    last_analysis_file: Optional[Path] = field(default=None)
    
    def save_state(self) -> None:
        """Saves current application state."""
        try:
            self.config.save()
            self._save_table_data()
            logger.info("Application state saved successfully")
        except Exception as e:
            logger.error(f"Error saving application state: {e}")
    
    def _save_table_data(self) -> None:
        """Saves current table data to temporary storage."""
        if not self.df_table.empty:
            try:
                temp_file = self.config.config_dir / 'last_analysis.pkl'
                self.df_table.to_pickle(temp_file)
                logger.debug("Table data saved to temporary storage")
            except Exception as e:
                logger.error(f"Error saving table data: {e}")
    
    def update_pagination(self) -> None:
        """Updates pagination information based on current data."""
        if self.df_table.empty:
            self.total_pages = 0
            self.current_page = 0
        else:
            self.total_pages = (len(self.df_table) - 1) // self.config.page_size + 1
            self.current_page = min(self.current_page, self.total_pages - 1)

class MRPGUI:
    """
    Main GUI class for the MRP Critical Items Analyzer application.
    Handles all user interface elements and interactions.
    """
    
    def __init__(self, root: tk.Tk):
        """
        Initialize the GUI application.
        
        Args:
            root: The root Tkinter window
        """
        self.root = root
        self._initialize_state()
        self._setup_window()
        self._create_variables()
        self._setup_bindings()
        self._build_ui()
        
    def _initialize_state(self) -> None:
        """Initialize application state and styling."""
        self.state = AppState()
        self.style = Style(self.state.config.theme)
        logger.info("Application state initialized")
        
    def _setup_window(self) -> None:
        """Configure main window properties."""
        self.root.title("MRP Critical Items Analyzer")
        self.root.geometry(
            f"{self.state.config.window_size[0]}x{self.state.config.window_size[1]}"
        )
        self.root.minsize(*self.state.config.min_window_size)
        
        # Set window icon if available
        icon_path = Path(__file__).parent / "assets" / "icon.ico"
        if icon_path.exists():
            self.root.iconbitmap(str(icon_path))
            
    def _create_variables(self) -> None:
        """Initialize Tkinter variables."""
        self.selected_file = tk.StringVar()
        self.sheet_name = tk.StringVar(value=self.state.config.default_sheet_name)
        self.filter_column = tk.StringVar()
        self.filter_value = tk.StringVar()
        self.qtd_min = tk.StringVar()
        self.qtd_max = tk.StringVar()
        
        # Comparison variables
        self.compare_before = None
        self.compare_after = None
        
    def _setup_bindings(self) -> None:
        """Setup keyboard shortcuts and event bindings."""
        # Window events
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        self.root.bind('<Configure>', self._on_window_configure)
        
        # Keyboard shortcuts
        self.root.bind('<Control-o>', lambda e: self._browse_file())
        self.root.bind('<Control-s>', lambda e: self._export_excel())
        self.root.bind('<Control-f>', lambda e: self._focus_filter())
        self.root.bind('<Control-r>', lambda e: self._run_analysis())
        self.root.bind('<Control-t>', lambda e: self._toggle_theme())
        
        logger.debug("Event bindings configured")
    
    def _setup_shortcuts(self):
        """Configura atalhos de teclado."""
        self.root.bind('<Control-o>', lambda e: self._browse_file())
        self.root.bind('<Control-s>', lambda e: self._export_excel())
        self.root.bind('<Control-f>', lambda e: self._focus_filter())
        self.root.bind('<Control-r>', lambda e: self._run_analysis())
        self.root.bind('<Control-t>', lambda e: self._toggle_theme())
        
    def _on_closing(self):
        """Handler para fechamento da janela."""
        try:
            self.state.config.window_size = (self.root.winfo_width(), self.root.winfo_height())
            self.state.save_state()
        finally:
            self.root.destroy()
            
    def _on_window_configure(self, event):
        """Handler para redimensionamento da janela."""
        if event.widget == self.root:
            pass

    def _toggle_theme(self):
        self.state.config.theme = "darkly" if self.state.config.theme == "flatly" else "flatly"
        self.style.theme_use(self.state.config.theme)
        self._log(f"Theme changed to: {self.state.config.theme}")

    def _build_ui(self):
        topbar = ttk.Frame(self.root)
        topbar.pack(fill=tk.X, pady=2)
        theme_btn = ttk.Button(topbar, text="Toggle Theme", command=self._toggle_theme)
        theme_btn.pack(side=tk.RIGHT, padx=10)
        ToolTip(theme_btn, text="Switch between light and dark mode (Ctrl+T)")
        about_btn = ttk.Button(topbar, text="About", command=self._show_about)
        about_btn.pack(side=tk.RIGHT, padx=10)
        self.root.bind('<Control-t>', lambda e: self._toggle_theme())

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.tab_analysis = ttk.Frame(self.notebook, padding=15)
        self.tab_table = ttk.Frame(self.notebook, padding=15)
        self.tab_compare = ttk.Frame(self.notebook, padding=15)

        self.notebook.add(self.tab_analysis, text="Analysis")
        self.notebook.add(self.tab_table, text="Table")
        self.notebook.add(self.tab_compare, text="Comparison")

        self._build_analysis_tab()
        self._build_table_tab()
        self._build_compare_tab()

    def _build_analysis_tab(self):
        form = ttk.Labelframe(self.tab_analysis, text="Run Analysis", padding=10)
        form.pack(pady=10, fill=tk.X)

        ttk.Label(form, text="Excel File:").grid(row=0, column=0, sticky=tk.E)
        entry_file = ttk.Entry(form, textvariable=self.selected_file, width=60)
        entry_file.grid(row=0, column=1, padx=5)
        btn_browse = ttk.Button(form, text="Browse", command=self._browse_file)
        btn_browse.grid(row=0, column=2)
        ToolTip(btn_browse, text="Select the Excel file to analyze")

        ttk.Label(form, text="Sheet Name:").grid(row=1, column=0, sticky=tk.E, pady=5)
        entry_sheet = ttk.Entry(form, textvariable=self.sheet_name, width=30)
        entry_sheet.grid(row=1, column=1, sticky=tk.W, pady=5)
        ToolTip(entry_sheet, text="Enter the worksheet name (e.g., Cálculo MRP)")

        btn_run = ttk.Button(form, text="Run Analysis", command=self._run_analysis, bootstyle="success")
        btn_run.grid(row=2, column=0, columnspan=3, pady=10)
        ToolTip(btn_run, text="Start the MRP analysis")

        self.progress = ttk.Progressbar(form, mode="indeterminate")
        self.progress.grid(row=3, column=0, columnspan=3, sticky=tk.EW)

        self.status_label = ttk.Label(form, text="", font=("Segoe UI", 10, "bold"))
        self.status_label.grid(row=4, column=0, columnspan=3, pady=5)

        ttk.Label(self.tab_analysis, text="Log:").pack(anchor=tk.W, padx=10)
        log_frame = ttk.Frame(self.tab_analysis)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD, bg="#f8f9fa", fg="#222")
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.config(yscrollcommand=scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _browse_file(self):
        file = filedialog.askopenfilename(filetypes=[("Excel Spreadsheets", "*.xlsx *.xls")])
        if file:
            self.selected_file.set(file)
            self._log(f"Selected file: {os.path.basename(file)}", "info")

    def _log(self, msg, level="info"):
        color = {"info": "#222", "success": "#155724", "error": "#721c24"}.get(level, "#222")
        self.log_text.insert(tk.END, f"{msg}\n", (level,))
        self.log_text.tag_config(level, foreground=color)
        self.log_text.see(tk.END)

    def _run_analysis(self) -> None:
        """
        Executes MRP analysis with enhanced feedback and robust error handling.
        Runs the analysis in a separate thread to prevent UI freezing.
        """
        try:
            file_path = Path(self.selected_file.get())
            sheet_name = self.sheet_name.get()
            
            self._validate_analysis_input(file_path, sheet_name)
            self._start_analysis_feedback()
            
            # Schedule analysis execution
            self.root.after(100, lambda: self._execute_analysis(file_path, sheet_name))
            
        except Exception as e:
            self._handle_analysis_error(str(e))
            
    def _validate_analysis_input(self, file_path: Path, sheet_name: str) -> None:
        """
        Validates input parameters for analysis.
        
        Args:
            file_path: Path to the input Excel file
            sheet_name: Name of the worksheet to analyze
            
        Raises:
            ValueError: If input parameters are invalid
            FileNotFoundError: If input file doesn't exist
        """
        if not file_path or str(file_path).strip() == "":
            raise ValueError("Please select a file to analyze.")
            
        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
            
        if not sheet_name or sheet_name.strip() == "":
            raise ValueError("Sheet name cannot be empty.")
            
        if not self._validate_excel_sheet(file_path, sheet_name):
            raise ValueError(f"Sheet '{sheet_name}' not found in the workbook.")
            
    def _validate_excel_sheet(self, file_path: Path, sheet_name: str) -> bool:
        """
        Validates that the specified sheet exists in the Excel file.
        
        Args:
            file_path: Path to the Excel file
            sheet_name: Name of the sheet to validate
            
        Returns:
            bool: True if sheet exists, False otherwise
        """
        try:
            return sheet_name in pd.ExcelFile(file_path).sheet_names
        except Exception as e:
            logger.error(f"Error validating sheet: {e}")
            return False
            
    def _start_analysis_feedback(self) -> None:
        """Configures visual feedback for analysis progress."""
        self.progress.start()
        self.status_label.config(
            text="Analyzing...",
            foreground="#007bff"
        )
        self._log("Starting analysis...", "info")
        self.root.update_idletasks()
        
    def _execute_analysis(self, file_path: Path, sheet_name: str) -> None:
        """
        Executes the MRP analysis with performance measurement.
        
        Args:
            file_path: Path to the input Excel file
            sheet_name: Name of the worksheet to analyze
        """
        try:
            start_time = time.time()
            
            output_file = (file_path.parent / 
                         f"itens_criticos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            
            # Execute analysis using MRPAnalyzer
            count, error, results = self.state.mrp_analyzer.analyze(
                str(file_path), 
                sheet_name, 
                str(output_file)
            )
            
            if error:
                raise Exception(error)
                
            self._handle_analysis_success(output_file, count, time.time() - start_time, results)
            
        except Exception as e:
            self._handle_analysis_error(str(e))
        finally:
            self.progress.stop()
            
    def _handle_analysis_success(self, output_file: Path, count: int, 
                               elapsed: float, results: pd.DataFrame) -> None:
        """
        Handles successful analysis completion.
        
        Args:
            output_file: Path to the output file
            count: Number of critical items found
            elapsed: Time taken for analysis
            results: DataFrame containing analysis results
        """
        elapsed = round(elapsed, 2)
        
        # Update UI with success information
        self._update_success_ui(elapsed, output_file)
        
        # Load results into table
        self._load_table(output_file)
        self.notebook.select(self.tab_table)
        
        # Show success message and offer to open file
        if self._show_success_dialog(count, elapsed, output_file):
            self._open_output_file(output_file)
            
        # Log success
        logger.info(
            f"Analysis completed successfully: {count} items found in {elapsed}s. "
            f"Output saved to {output_file}"
        )
        
    def _update_success_ui(self, elapsed: float, output_file: Path) -> None:
        """Updates UI elements after successful analysis."""
        self._log(
            f"Analysis completed in {elapsed}s", 
            "success"
        )
        self.status_label.config(
            text=f"Output file: {output_file.name}",
            foreground="#28a745"
        )
        
    def _show_success_dialog(self, count: int, elapsed: float, 
                           output_file: Path) -> bool:
        """
        Shows success dialog and returns whether to open the output file.
        
        Returns:
            bool: True if user wants to open the output file
        """
        message = (
            f"Analysis completed successfully!\n\n"
            f"Time: {elapsed}s\n"
            f"Critical Items: {count}\n"
            f"Output: {output_file.name}\n\n"
            "Would you like to open the generated file?"
        )
        return messagebox.askyesno("Success", message)
        
    def _open_output_file(self, file_path: Path) -> None:
        """
        Opens the output file in the default application.
        
        Args:
            file_path: Path to the file to open
        """
        try:
            webbrowser.open(str(file_path))
            logger.info(f"Opened output file: {file_path}")
        except Exception as e:
            logger.error(f"Error opening file: {e}")
            messagebox.showerror(
                "Error",
                f"Could not open the file: {str(e)}"
            )

    def _handle_analysis_error(self, error: str) -> None:
        """
        Handles analysis errors and updates UI accordingly.
        
        Args:
            error: Error message to display
        """
        # Log error
        logger.error(f"Analysis error: {error}")
        
        # Update UI
        self._log(f"Error during analysis: {error}", "error")
        self.status_label.config(
            text="Analysis failed",
            foreground="#dc3545"
        )
        
        # Show error dialog
        messagebox.showerror(
            "Analysis Error",
            f"An error occurred during analysis:\n\n{error}"
        )
        
        # Reset progress
        self.progress.stop()
        self.root.update_idletasks()

    def _validate_sheet(self, file, sheet):
        try:
            return sheet in pd.ExcelFile(file).sheet_names
        except Exception as e:
            self._log(f"Error validating sheet: {e}", "error")
            return False

    def _build_table_tab(self):
        filter_frame = ttk.Labelframe(self.tab_table, text="Filter & Export", padding=10)
        filter_frame.pack(fill=tk.X, pady=5)

        self.filter_column = tk.StringVar()
        self.filter_value = tk.StringVar()
        self.qtd_min = tk.StringVar()
        self.qtd_max = tk.StringVar()

        self.column_box = ttk.Combobox(filter_frame, textvariable=self.filter_column, state="readonly", width=30)
        self.column_box.pack(side=tk.LEFT, padx=5)
        ToolTip(self.column_box, text="Select column to filter")
        entry_filter = ttk.Entry(filter_frame, textvariable=self.filter_value, width=30)
        entry_filter.pack(side=tk.LEFT)
        ToolTip(entry_filter, text="Enter value to filter")

        ttk.Label(filter_frame, text="Min Qty:").pack(side=tk.LEFT, padx=2)
        entry_min = ttk.Entry(filter_frame, textvariable=self.qtd_min, width=6)
        entry_min.pack(side=tk.LEFT)
        ToolTip(entry_min, text="Minimum quantity to request")

        ttk.Label(filter_frame, text="Max Qty:").pack(side=tk.LEFT, padx=2)
        entry_max = ttk.Entry(filter_frame, textvariable=self.qtd_max, width=6)
        entry_max.pack(side=tk.LEFT)
        ToolTip(entry_max, text="Maximum quantity to request")

        btn_filter = ttk.Button(filter_frame, text="Apply Filter", command=self._apply_filter)
        btn_filter.pack(side=tk.LEFT, padx=5)
        ToolTip(btn_filter, text="Apply filter to table")
        btn_reload = ttk.Button(filter_frame, text="Reload", command=self._load_table)
        btn_reload.pack(side=tk.LEFT)
        ToolTip(btn_reload, text="Reload table from file")

        btn_export_excel = ttk.Button(filter_frame, text="Export Excel", command=self._export_excel)
        btn_export_excel.pack(side=tk.RIGHT, padx=5)
        ToolTip(btn_export_excel, text="Export table to Excel file")
        btn_export_csv = ttk.Button(filter_frame, text="Export CSV", command=self._export_csv)
        btn_export_csv.pack(side=tk.RIGHT)
        ToolTip(btn_export_csv, text="Export table to CSV file")

        self.tree = ttk.Treeview(self.tab_table, show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True)

        nav_frame = ttk.Frame(self.tab_table)
        nav_frame.pack(fill=tk.X, pady=10)

        self.stats_label = ttk.Label(nav_frame, text="")
        self.stats_label.pack(side=tk.LEFT, padx=10)

        btn_frame = ttk.Frame(nav_frame)
        btn_frame.pack(side=tk.RIGHT)
        btn_prev = ttk.Button(btn_frame, text="Previous", command=self._prev_page)
        btn_prev.pack(side=tk.LEFT, padx=5)
        ToolTip(btn_prev, text="Previous page")
        btn_next = ttk.Button(btn_frame, text="Next", command=self._next_page)
        btn_next.pack(side=tk.LEFT)
        ToolTip(btn_next, text="Next page")

    def _load_table(self, path: Optional[Path] = None) -> None:
        """
        Loads and displays table data from an Excel file.
        
        Args:
            path: Optional path to the Excel file. If not provided,
                  uses the last analysis file.
        """
        try:
            file_path = path or Path(self.selected_file.get()).parent / "itens_criticos.xlsx"
            
            # Load data with optimized settings
            self.state.df_table = pd.read_excel(
                file_path,
                dtype={
                    'CÓD': str,
                    'QUANTIDADE A SOLICITAR': 'Int64',
                    'ESTOQUE DISPONÍVEL': 'Int64'
                }
            )
            
            # Update UI elements
            self.column_box['values'] = list(self.state.df_table.columns)
            self.state.current_page = 0
            self.state.update_pagination()
            
            # Render table and update statistics
            self._render_table()
            logger.info(f"Table loaded successfully from {file_path}")
            
        except Exception as e:
            logger.error(f"Error loading table: {e}")
            self._log(f"Error loading table: {str(e)}", "error")
            messagebox.showerror("Error", f"Failed to load table: {str(e)}")

    def _render_table(self) -> None:
        """
        Renders the table with efficient pagination and caching.
        Updates statistics and pagination information.
        """
        # Clear existing data
        self.tree.delete(*self.tree.get_children())
        df = self.state.df_table
        
        if df.empty:
            self._update_stats({
                'total': 0,
                'soma': 0,
                'media': 0,
                'top_forn': '-'
            })
            return
        
        # Update statistics if needed
        if not hasattr(self, '_stats_cache') or self.state.filter_applied:
            self._stats_cache = self._calculate_statistics(df)
            self.state.filter_applied = False
        
        if not self.tree["columns"]:
            self.tree["columns"] = list(df.columns)
            for col in df.columns:
                self.tree.heading(col, text=col, command=lambda c=col: self._sort_column(c))
                self.tree.column(col, width=120, anchor="center")
                
        # Get current page data
        start_idx = self.state.current_page * self.state.config.page_size
        end_idx = start_idx + self.state.config.page_size
        current_page = df.iloc[start_idx:end_idx]
        
        # Render rows with alternating colors
        for i, (_, row) in enumerate(current_page.iterrows()):
            tags = ('oddrow',) if i % 2 else ('evenrow',)
            self.tree.insert("", tk.END, values=list(row), tags=tags)

        # Update statistics display
        self._update_display_statistics()
        
    def _calculate_statistics(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        Calculates table statistics.
        
        Args:
            df: DataFrame to analyze
            
        Returns:
            Dict containing calculated statistics
        """
        try:
            stats = {
                'total': len(df),
                'soma': 0,
                'media': 0,
                'top_forn': '-'
            }
            
            if "QUANTIDADE A SOLICITAR" in df.columns:
                qty_series = df["QUANTIDADE A SOLICITAR"]
                stats.update({
                    'soma': int(qty_series.sum()),
                    'media': round(qty_series.mean(), 2)
                })
                
            if "FORNECEDOR PRINCIPAL" in df.columns:
                stats['top_forn'] = df["FORNECEDOR PRINCIPAL"].value_counts().idxmax()
                
            return stats
            
        except Exception as e:
            logger.error(f"Error calculating statistics: {e}")
            return {
                'total': 0,
                'soma': 0,
                'media': 0,
                'top_forn': 'Error'
            }
            
    def _update_display_statistics(self) -> None:
        """Updates the statistics display in the UI."""
        stats = self._stats_cache
        
        # Update statistics label
        self.stats_label.config(
            text=(f"Total Items: {stats['total']} | "
                  f"Total Quantity: {stats['soma']} | "
                  f"Average: {stats['media']} | "
                  f"Top Supplier: {stats['top_forn']}")
        )
        
        # Update pagination label
        self.page_label.config(
            text=f"Page {self.state.current_page + 1} of {self.state.total_pages}"
        )

    def _apply_filter(self):
        df = self.df_table.copy()
        col = self.filter_column.get()
        val = self.filter_value.get().strip().lower()
        min_qtd = self.qtd_min.get()
        max_qtd = self.qtd_max.get()

        if col and val:
            df = df[df[col].astype(str).str.lower().str.contains(val)]

        if "QUANTIDADE A SOLICITAR" in df.columns:
            if min_qtd.isdigit():
                df = df[df["QUANTIDADE A SOLICITAR"] >= int(min_qtd)]
            if max_qtd.isdigit():
                df = df[df["QUANTIDADE A SOLICITAR"] <= int(max_qtd)]

        self.df_table = df
        self.current_page = 0
        self._render_table()

    def _sort_column(self, col):
        self.df_table.sort_values(by=col, ascending=True, inplace=True, ignore_index=True)
        self.current_page = 0
        self._render_table()

    def _prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self._render_table()

    def _next_page(self):
        if (self.current_page + 1) * self.page_size < len(self.df_table):
            self.current_page += 1
            self._render_table()

    def _export_csv(self):
        file = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if file:
            self.df_table.to_csv(file, index=False)
            self._log(f"CSV saved: {file}", "success")
            messagebox.showinfo("Export", f"CSV file saved: {file}")

    def _export_excel(self):
        file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if file:
            self.df_table.to_excel(file, index=False)
            self._log(f"Excel saved: {file}", "success")
            messagebox.showinfo("Export", f"Excel file saved: {file}")

    def _build_compare_tab(self):
        frame = ttk.Labelframe(self.tab_compare, text="Compare Analyses", padding=10)
        frame.pack(pady=10, fill=tk.X)

        btn_before = ttk.Button(frame, text="Select Previous Analysis", command=self._load_before)
        btn_before.pack(side=tk.LEFT, padx=5)
        ToolTip(btn_before, text="Select the previous analysis Excel file")
        btn_after = ttk.Button(frame, text="Select Current Analysis", command=self._load_after)
        btn_after.pack(side=tk.LEFT, padx=5)
        ToolTip(btn_after, text="Select the current analysis Excel file")
        btn_compare = ttk.Button(frame, text="Compare", command=self._compare_files, bootstyle="info")
        btn_compare.pack(side=tk.LEFT, padx=10)
        ToolTip(btn_compare, text="Compare the two analyses")

        self.compare_tree = ttk.Treeview(self.tab_compare, show="headings")
        self.compare_tree.pack(fill=tk.BOTH, expand=True)

    def _load_before(self):
        file = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if file:
            self.compare_before = pd.read_excel(file)
            self._log(f"Previous analysis loaded: {os.path.basename(file)}", "info")

    def _load_after(self):
        file = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if file:
            self.compare_after = pd.read_excel(file)
            self._log(f"Current analysis loaded: {os.path.basename(file)}", "info")

    def _compare_files(self):
        if self.compare_before is None or self.compare_after is None:
            messagebox.showwarning("Missing File", "Load both analyses to compare.")
            self._log("Both analyses must be loaded for comparison.", "error")
            return
        if self.compare_before.empty or self.compare_after.empty:
            messagebox.showwarning("Empty Analysis", "One or both analyses are empty.")
            self._log("One or both analyses are empty.", "error")
            return
        before = self.compare_before.set_index("CÓD")
        after = self.compare_after.set_index("CÓD")

        all_codes = sorted(set(before.index) | set(after.index))
        result = []

        for code in all_codes:
            row = {"CÓD": code}
            row["DESCRIÇÃO"] = after.at[code, "DESCRIÇÃOPROMOB"] if code in after.index else before.at[code, "DESCRIÇÃOPROMOB"]
            row["FORNECEDOR"] = after.at[code, "FORNECEDOR PRINCIPAL"] if code in after.index else before.at[code, "FORNECEDOR PRINCIPAL"]
            q_ant = before.at[code, "QUANTIDADE A SOLICITAR"] if code in before.index else 0
            q_atu = after.at[code, "QUANTIDADE A SOLICITAR"] if code in after.index else 0
            row["ANTERIOR"] = q_ant
            row["ATUAL"] = q_atu
            row["DIFERENÇA"] = q_atu - q_ant

            if code not in before.index:
                row["STATUS"] = "New"
            elif code not in after.index:
                row["STATUS"] = "Removed"
            elif q_ant != q_atu:
                row["STATUS"] = "Changed"
            else:
                row["STATUS"] = "Unchanged"

            result.append(row)

        df = pd.DataFrame(result)
        self.compare_tree.delete(*self.compare_tree.get_children())
        self.compare_tree["columns"] = list(df.columns)

        status_colors = {
            "New": "#d4edda",
            "Removed": "#f8d7da",
            "Changed": "#fff3cd",
            "Unchanged": "#f9f9f9"
        }
        for col in df.columns:
            self.compare_tree.heading(col, text=col)
            self.compare_tree.column(col, width=120, anchor="center")
        for _, row in df.iterrows():
            tag = row["STATUS"]
            self.compare_tree.insert("", tk.END, values=list(row), tags=(tag,))
        for status, color in status_colors.items():
            self.compare_tree.tag_configure(status, background=color)
        for col in df.columns:
            max_len = max([len(str(x)) for x in df[col].values] + [len(col)])
            self.compare_tree.column(col, width=min(200, max(80, max_len * 10)))
        for col in df.columns:
            self.compare_tree.heading(col, text=col, command=lambda c=col: self._sort_compare_column(c))
            
    def _show_about(self):
        messagebox.showinfo(
            "About",
            "MRP Critical Items Analyzer\n\nDeveloped by Humberto Rodrigues.\nModern UI, color feedback, and Excel/CSV export.\n2025"
        )

def main():
    root = tk.Tk()
    app = MRPGUI(root)
    root.mainloop()

def set_style():
    style = Style("flatly")
    style.configure("TButton", padding=5, relief="flat")
    style.configure("TLabel", padding=5)
    style.configure("TEntry", padding=5)
    style.configure("TFrame", padding=10)
    return style

if __name__ == "__main__":
    main()
