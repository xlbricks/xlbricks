"""
    Author: julij.jegorov
    Date: 15/02/2026
    Description: PyQt UI for editing xlbricks.json configuration;
                 loaded from Excel (e.g. XLBricks config editor / Wizard).
"""

import json
import os
import os.path as osp
from PyQt5 import QtCore
from PyQt5.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QFormLayout,
    QLineEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QFileDialog,
    QMessageBox,
    QGroupBox,
    QLabel,
    QAbstractItemView,
)
from PyQt5.QtCore import Qt


def get_default_config_path():
    """Get the default location of the xlbricks configuration file.
    
    Returns the path to xlbricks.json in the package directory.
    """
    pkg_dir = osp.dirname(osp.dirname(osp.abspath(__file__)))
    return osp.join(pkg_dir, 'xlbricks.json')


def load_config(path):
    """Load config from JSON file. Returns dict or None on error."""
    if not path or not osp.isfile(path):
        return _default_config()
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return _normalize_config(data)
    except Exception:
        return _default_config()


def _default_config():
    """Return default configuration structure.
    
    Provides empty values for all required config keys.
    """
    return {
        'APPS_PATH': '',
        'PYTHONPATH': '',
        'CONTEXT': {},
    }


def _normalize_config(data):
    """Ensure all required configuration keys exist.
    
    Fills in missing keys with defaults for backward compatibility.
    """
    out = _default_config()
    # APPS_PATH: path to xlbricks applications (e.g. .../technology/apps)
    out['APPS_PATH'] = data.get('APPS_PATH', data.get('INTERPRETER', ''))
    out['PYTHONPATH'] = data.get('PYTHONPATH', '')
    out['CONTEXT'] = dict(data.get('CONTEXT', {}))
    return out


def save_config(path, data):
    """Write configuration to a JSON file.
    
    Saves with proper formatting and UTF-8 encoding.
    """
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)


class ConfigEditorDialog(QDialog):
    """Interactive dialog for editing XLBricks settings.
    
    Provides a user-friendly interface for managing paths, Python paths, and contexts.
    """

    def __init__(self, config_path=None, parent=None):
        """Initialize the config editor dialog.
        
        Opens the specified config file or uses the default xlbricks.json.
        """
        super(ConfigEditorDialog, self).__init__(parent)
        self._config_path = config_path or get_default_config_path()
        self.setWindowTitle('XLBricks Settings')
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.Window)
        self.setMinimumSize(650, 600)
        self.resize(650, 500)
        self._build_ui()
        self._load_into_ui()

    def _build_ui(self):
        """Construct the dialog UI with all input fields and buttons."""
        layout = QVBoxLayout(self)

        # --- XLBricks applications path ---
        grp_apps = QGroupBox('XLBricks applications path')
        grp_apps.setToolTip("Folder containing xlbricks applications. Files will be access with 'Load Apps'")
        fl_apps = QFormLayout(grp_apps)
        self._apps_path_edit = QLineEdit()
        self._apps_path_edit.setPlaceholderText('Path to apps')
        self._apps_path_edit.setMinimumWidth(320)
        btn_browse_apps = QPushButton('Browse...')
        btn_browse_apps.setMaximumWidth(90)
        btn_browse_apps.clicked.connect(self._browse_apps_path)
        row = QHBoxLayout()
        row.addWidget(self._apps_path_edit)
        row.addWidget(btn_browse_apps)
        fl_apps.addRow('Path:', row)
        layout.addWidget(grp_apps)

        # --- PYTHONPATH ---
        grp_path = QGroupBox('PYTHONPATH')
        grp_path.setToolTip('Paths added to Python when running from Excel. One path per row.')
        path_layout = QVBoxLayout(grp_path)
        self._path_table = QTableWidget()
        self._path_table.setColumnCount(1)
        self._path_table.setHorizontalHeaderLabels(['Path'])
        self._path_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self._path_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._path_table.setMinimumHeight(100)
        path_layout.addWidget(self._path_table)
        path_btn_row = QHBoxLayout()
        btn_add_path = QPushButton('Add path')
        btn_add_path_browse = QPushButton('Browse...')
        btn_remove_path = QPushButton('Remove selected')
        btn_add_path.clicked.connect(self._add_path_row)
        btn_add_path_browse.clicked.connect(self._browse_path_row)
        btn_remove_path.clicked.connect(self._remove_path_row)
        path_btn_row.addWidget(btn_add_path)
        path_btn_row.addWidget(btn_add_path_browse)
        path_btn_row.addWidget(btn_remove_path)
        path_btn_row.addStretch()
        path_layout.addLayout(path_btn_row)
        layout.addWidget(grp_path)

        # --- CONTEXT ---
        grp_context = QGroupBox('Context')
        grp_context.setToolTip('Context name → module path. Used to resolve context objects.')
        ctx_layout = QVBoxLayout(grp_context)
        self._context_table = QTableWidget()
        self._context_table.setColumnCount(2)
        self._context_table.setHorizontalHeaderLabels(['Context name', 'Module path'])
        self._context_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
        self._context_table.setColumnWidth(0, 180)
        self._context_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self._context_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._context_table.setMinimumHeight(140)
        ctx_layout.addWidget(self._context_table)
        btn_row = QHBoxLayout()
        btn_add_ctx = QPushButton('Add context')
        btn_remove_ctx = QPushButton('Remove selected')
        btn_add_ctx.clicked.connect(self._add_context_row)
        btn_remove_ctx.clicked.connect(self._remove_context_row)
        btn_row.addWidget(btn_add_ctx)
        btn_row.addWidget(btn_remove_ctx)
        btn_row.addStretch()
        ctx_layout.addLayout(btn_row)
        layout.addWidget(grp_context)

        # --- Config file path (read-only) ---
        self._path_label = QLabel(self._config_path)
        self._path_label.setStyleSheet('color: gray; font-size: 11px;')
        self._path_label.setWordWrap(True)
        layout.addWidget(self._path_label)

        # --- Buttons ---
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self._save_btn = QPushButton('Save')
        self._save_btn.setDefault(True)
        self._save_btn.setMinimumWidth(90)
        self._cancel_btn = QPushButton('Cancel')
        self._cancel_btn.setMinimumWidth(90)
        self._save_btn.clicked.connect(self._save)
        self._cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(self._save_btn)
        btn_layout.addWidget(self._cancel_btn)
        layout.addLayout(btn_layout)

    def _browse_apps_path(self):
        """Open a folder picker for selecting the apps directory."""
        path = QFileDialog.getExistingDirectory(
            self,
            'Select XLBricks applications folder',
            self._apps_path_edit.text() or os.path.expanduser('~'),
        )
        if path:
            self._apps_path_edit.setText(path)

    def _add_context_row(self):
        """Add a new empty row to the context mapping table."""
        row = self._context_table.rowCount()
        self._context_table.insertRow(row)
        self._context_table.setItem(row, 0, QTableWidgetItem(''))
        self._context_table.setItem(row, 1, QTableWidgetItem(''))

    def _remove_context_row(self):
        """Delete the selected row from the context table."""
        row = self._context_table.currentRow()
        if row >= 0:
            self._context_table.removeRow(row)

    def _add_path_row(self):
        """Add a new empty row to the PYTHONPATH table."""
        row = self._path_table.rowCount()
        self._path_table.insertRow(row)
        self._path_table.setItem(row, 0, QTableWidgetItem(''))

    def _browse_path_row(self):
        """Open a folder picker and add the selected path to PYTHONPATH."""
        path = QFileDialog.getExistingDirectory(self, 'Select folder to add to PYTHONPATH')
        if path:
            row = self._path_table.rowCount()
            self._path_table.insertRow(row)
            self._path_table.setItem(row, 0, QTableWidgetItem(path))

    def _remove_path_row(self):
        """Delete the selected row from the PYTHONPATH table."""
        row = self._path_table.currentRow()
        if row >= 0:
            self._path_table.removeRow(row)

    def _load_into_ui(self):
        """Populate UI fields with values from the config file."""
        data = load_config(self._config_path)
        self._apps_path_edit.setText(data.get('APPS_PATH', ''))
        path_str = data.get('PYTHONPATH', '')
        # Support both newline and semicolon separation
        if ';' in path_str and '\n' not in path_str:
            paths = [p.strip() for p in path_str.split(';') if p.strip()]
        else:
            paths = [p.strip() for p in path_str.replace(';', '\n').splitlines() if p.strip()]
        self._path_table.setRowCount(0)
        for p in paths:
            row = self._path_table.rowCount()
            self._path_table.insertRow(row)
            self._path_table.setItem(row, 0, QTableWidgetItem(p))
        ctx = data.get('CONTEXT', {})
        self._context_table.setRowCount(0)
        for name, mod in ctx.items():
            row = self._context_table.rowCount()
            self._context_table.insertRow(row)
            self._context_table.setItem(row, 0, QTableWidgetItem(name))
            self._context_table.setItem(row, 1, QTableWidgetItem(mod))

    def _collect_from_ui(self):
        """Gather all values from UI fields into a config dictionary.
        
        Prepares data for saving to the config file.
        """
        paths = []
        for r in range(self._path_table.rowCount()):
            item = self._path_table.item(r, 0)
            p = (item.text() if item else '').strip()
            if p:
                paths.append(p)
        path_sep = os.pathsep
        ctx = {}
        for r in range(self._context_table.rowCount()):
            name_item = self._context_table.item(r, 0)
            mod_item = self._context_table.item(r, 1)
            name = (name_item.text() if name_item else '').strip()
            mod = (mod_item.text() if mod_item else '').strip()
            if name:
                ctx[name] = mod
        return {
            'APPS_PATH': self._apps_path_edit.text().strip(),
            'PYTHONPATH': path_sep.join(paths),
            'CONTEXT': ctx,
        }

    def _save(self):
        try:
            # Merge UI data into existing config so we don't remove keys we don't edit (e.g. INTERPRETER)
            existing = {}
            if osp.isfile(self._config_path):
                try:
                    with open(self._config_path, 'r', encoding='utf-8') as f:
                        existing = json.load(f)
                except Exception:
                    pass
            data = self._collect_from_ui()
            for key, value in data.items():
                existing[key] = value
            dir_path = osp.dirname(self._config_path)
            if dir_path and not osp.isdir(dir_path):
                os.makedirs(dir_path, exist_ok=True)
            save_config(self._config_path, existing)
            from os.path import basename
            msg_box = QMessageBox(QMessageBox.Information, 'Saved', f'Config saved to:\n{basename(self._config_path)}', parent=self)
            msg_box.exec_()
            self.accept()
        except Exception as e:
            QMessageBox.critical(
                self,
                'Save failed',
                'Could not save config:\n' + str(e),
            )


def show_config_editor(config_path=None, parent=None):
    """Display the configuration editor dialog.
    
    Returns True if user saved changes, False if cancelled.
    """
    dlg = ConfigEditorDialog(config_path=config_path, parent=parent)
    return dlg.exec_() == QDialog.Accepted
