#!/usr/bin/env python3
"""Shiori - GUI application to add bookmarks to PDF files."""

import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple

import fitz  # PyMuPDF
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QKeySequence, QShortcut
from PyQt6.QtWidgets import (
    QApplication, QDialog, QDialogButtonBox, QFileDialog,
    QGridLayout, QGroupBox, QHBoxLayout, QLabel, QLineEdit,
    QMainWindow, QMessageBox, QPushButton, QSpinBox,
    QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget,
)

# ── Stylesheet ────────────────────────────────────────────────────────────────

_STYLE = """
QLineEdit {
    background-color: #DFF0FF;
    border: 1px solid #85C1E9;
    border-radius: 3px;
    padding: 3px 6px;
}
QSpinBox {
    background-color: #DFF0FF;
    border: 1px solid #85C1E9;
    border-radius: 3px;
    padding: 2px 4px;
}
QPushButton {
    background-color: #5D6D7E;
    color: white;
    border: none;
    border-radius: 4px;
    padding: 5px 12px;
    min-width: 52px;
}
QPushButton:hover   { background-color: #7F8C8D; }
QPushButton:pressed { background-color: #424949; }

QPushButton#primary {
    background-color: #1F618D;
    font-weight: bold;
    padding: 6px 18px;
}
QPushButton#primary:hover   { background-color: #2E86C1; }
QPushButton#primary:pressed { background-color: #154360; }
"""

# ── Text normalization ────────────────────────────────────────────────────────

def _to_half(s: str) -> str:
    """Convert full-width ASCII characters (U+FF01–U+FF5E) to half-width."""
    out = []
    for c in s:
        cp = ord(c)
        out.append(chr(cp - 0xFEE0) if 0xFF01 <= cp <= 0xFF5E else c)
    return "".join(out)


# ── Heading patterns ──────────────────────────────────────────────────────────
# Applied after _to_half() normalization.  First match wins; most specific first.

_HEADING_PATTERNS: List[Tuple[int, re.Pattern]] = [
    # Japanese chapter system
    (3, re.compile(r"^第\d+\.\d+\.\d+節")),
    (2, re.compile(r"^第\d+\.\d+節")),
    (1, re.compile(r"^第(?:\d+|[一二三四五六七八九十百千万零〇]+)章")),
    # Numeric system (3-level before 2-level before 1-level)
    (3, re.compile(r"^\d+\.\d+\.\d+")),
    (2, re.compile(r"^\d+\.\d+\.?(?:[ \t　]|$)")),
    (1, re.compile(r"^\d+\.(?:[ \t　]|$)")),
    # Parenthesis system
    (3, re.compile(r"^\(\d+\)(?:[ \t　]|$)")),
    (4, re.compile(r"^\([ア-ン]+\)(?:[ \t　]|$)")),
    (5, re.compile(r"^\([a-zA-Z]\)(?:[ \t　]|$)")),
]

# ── Figure / table patterns ───────────────────────────────────────────────────
# kind: "figure" | "table"

_FIG_TABLE_PATTERNS: List[Tuple[str, re.Pattern]] = [
    ("figure", re.compile(r"^図\s*[0-9]+(?:[.\-・][0-9]+)*")),
    ("figure", re.compile(r"^Figure\s+[0-9]+", re.IGNORECASE)),
    ("table",  re.compile(r"^表\s*[0-9]+(?:[.\-・][0-9]+)*")),
    ("table",  re.compile(r"^Table\s+[0-9]+",  re.IGNORECASE)),
]

# Particles / functional words that — when found immediately after a figure/table number —
# mark the line as body text referencing the figure, not a caption.
#
#   compound particles  : では  には  中の  中に  中で  より  から
#   single particles    : が  を  に  の
#   formal connectives  : において  について  による  として  からは  …
#                         (all start with one of the single particles above)
#
# NOTE: 「と」と「で」は単体だと誤検知リスクがあるため除外。
_BODY_AFTER_FIG = re.compile(
    r"^(?:では|には|より|から|中[のにで]|[がをに]|の)"
)


def _match_heading(line: str) -> Optional[int]:
    normed = _to_half(line.strip())
    for level, pat in _HEADING_PATTERNS:
        if pat.match(normed):
            return level
    return None


def _match_fig_table(line: str) -> Optional[str]:
    normed = _to_half(line.strip())
    for kind, pat in _FIG_TABLE_PATTERNS:
        m = pat.match(normed)
        if not m:
            continue
        # Check what immediately follows the number (after stripping leading spaces).
        # If it is a Japanese particle/functional word, this is a body text reference.
        after = normed[m.end():].lstrip()
        if after and _BODY_AFTER_FIG.match(after):
            continue
        return kind
    return None


# ── Data model ────────────────────────────────────────────────────────────────

@dataclass
class TitleEntry:
    raw_level: int    # pattern level for headings; 0 for figure/table
    disp_level: int   # display level after gap-promotion
    text: str
    page: int         # 0-indexed
    y: float
    kind: str = "heading"   # "heading" | "figure" | "table"


def promote_levels(entries: List[TitleEntry]) -> List[TitleEntry]:
    """Renumber heading disp_levels to 1, 2, 3, … with no gaps. Leaves figure/table untouched."""
    headings = [e for e in entries if e.kind == "heading"]
    if not headings:
        return entries
    used = sorted({e.raw_level for e in headings})
    mapping = {raw: i + 1 for i, raw in enumerate(used)}
    for e in headings:
        e.disp_level = mapping[e.raw_level]
    return entries


# ── PDF extraction ────────────────────────────────────────────────────────────

def extract_titles(doc: fitz.Document, start_page: int = 0) -> List[TitleEntry]:
    """Extract heading + figure/table entries, skipping pages before start_page (0-indexed)."""
    entries: List[TitleEntry] = []
    for page_idx, page in enumerate(doc):
        if page_idx < start_page:
            continue
        for block in page.get_text("blocks"):
            if block[6] != 0:
                continue
            raw: str = block[4]
            y: float = block[1]
            lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
            if not lines:
                continue

            # Heading?
            level = _match_heading(lines[0])
            if level is not None:
                title = lines[0]
                if len(lines) >= 2 and len(_to_half(title)) <= 15:
                    title = title + " " + lines[1]
                entries.append(TitleEntry(
                    raw_level=level, disp_level=level,
                    text=title.strip(), page=page_idx, y=y,
                    kind="heading",
                ))
                continue

            # Figure or table?
            kind = _match_fig_table(lines[0])
            if kind is not None:
                title = lines[0]
                if len(lines) >= 2 and len(_to_half(title)) <= 20:
                    title = title + " " + lines[1]
                entries.append(TitleEntry(
                    raw_level=0, disp_level=2,   # shown under the separator (level 2)
                    text=title.strip(), page=page_idx, y=y,
                    kind=kind,
                ))

    return entries


# ── TOC building ──────────────────────────────────────────────────────────────

def _fix_toc_levels(toc: list) -> list:
    """Clamp levels: first must be 1, no entry may jump more than +1."""
    if not toc:
        return toc
    toc[0][0] = 1
    for i in range(1, len(toc)):
        toc[i][0] = min(toc[i][0], toc[i - 1][0] + 1)
    return toc


def build_toc(entries: List[TitleEntry]) -> list:
    """Heading bookmarks first, then a separator + figure/table bookmarks."""
    headings  = [e for e in entries if e.kind == "heading"]
    figtables = [e for e in entries if e.kind in ("figure", "table")]

    toc: list = [[e.disp_level, e.text, e.page + 1] for e in headings]

    if figtables:
        toc.append([1, "── 図表のしおり ──", figtables[0].page + 1])
        for e in figtables:
            toc.append([2, e.text, e.page + 1])

    _fix_toc_levels(toc)
    return toc


def save_pdf_with_bookmarks(
    input_path: str, output_path: str, entries: List[TitleEntry]
) -> int:
    doc = fitz.open(input_path)
    toc = build_toc(entries)
    doc.set_toc(toc)
    doc.save(output_path, garbage=4, deflate=True)
    doc.close()

    verify = fitz.open(output_path)
    count = len(verify.get_toc())
    verify.close()
    return count


def default_output(input_path: str) -> str:
    p = Path(input_path)
    return str(p.parent / (p.stem + "_out.pdf"))


# ── Tree colors ───────────────────────────────────────────────────────────────

_COL_FIGURE = QColor("#154360")   # dark blue  – figure rows
_COL_TABLE  = QColor("#145A32")   # dark green – table rows
_COL_SEP_FG = QColor("#777777")   # gray text  – separator row
_COL_SEP_BG = QColor("#EEEEEE")   # gray bg    – separator row


# ── Edit dialog ───────────────────────────────────────────────────────────────

class EditDialog(QDialog):
    def __init__(self, entry: TitleEntry, parent=None):
        super().__init__(parent)
        self.setWindowTitle("テキスト編集")
        self.setFixedWidth(520)
        self.setWindowModality(Qt.WindowModality.ApplicationModal)

        layout = QVBoxLayout(self)
        kind_tag = {"heading": f"H{entry.disp_level}", "figure": "図", "table": "表"}
        layout.addWidget(QLabel(f"{kind_tag.get(entry.kind, '?')} / ページ {entry.page + 1}"))

        self._edit = QLineEdit(entry.text)
        layout.addWidget(self._edit)

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save |
            QDialogButtonBox.StandardButton.Cancel
        )
        btns.button(QDialogButtonBox.StandardButton.Save).setText("保存")
        btns.button(QDialogButtonBox.StandardButton.Cancel).setText("キャンセル")
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

        self._edit.setFocus()
        self._edit.selectAll()

    def get_text(self) -> str:
        return self._edit.text().strip()


# ── Main window ───────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Shiori - PDF しおり追加")
        self.resize(900, 640)
        self.setMinimumSize(700, 440)
        self._entries: List[TitleEntry] = []
        self._setup_ui()

    # ── Layout ────────────────────────────────────────────────────────────────

    def _setup_ui(self):
        self.setStyleSheet(_STYLE)

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setSpacing(6)
        root.setContentsMargins(12, 8, 12, 8)

        # ── File group ────────────────────────────────────────────────────────
        file_group = QGroupBox("ファイル")
        grid = QGridLayout(file_group)
        grid.setColumnStretch(1, 1)

        self._in_edit  = QLineEdit()
        self._out_edit = QLineEdit()
        self._in_edit.setPlaceholderText("入力 PDF のパスを入力するか「参照」で選択")
        self._out_edit.setPlaceholderText("出力 PDF の保存先（空欄で自動設定）")

        for row, label_text, edit, slot in (
            (0, "入力 PDF:", self._in_edit,  self._pick_input),
            (1, "出力 PDF:", self._out_edit, self._pick_output),
        ):
            grid.addWidget(QLabel(label_text), row, 0)
            grid.addWidget(edit, row, 1)
            btn = QPushButton("参照")
            btn.setFixedWidth(60)
            btn.clicked.connect(slot)
            grid.addWidget(btn, row, 2)

        root.addWidget(file_group)

        # ── Detect row ────────────────────────────────────────────────────────
        ctrl = QHBoxLayout()

        detect_btn = QPushButton("タイトルを検出")
        detect_btn.setObjectName("primary")
        detect_btn.clicked.connect(self._detect)
        ctrl.addWidget(detect_btn)

        ctrl.addSpacing(20)
        ctrl.addWidget(QLabel("本文開始ページ:"))

        self._start_spin = QSpinBox()
        self._start_spin.setMinimum(1)
        self._start_spin.setMaximum(9999)
        self._start_spin.setValue(1)
        self._start_spin.setFixedWidth(64)
        self._start_spin.setToolTip("表紙・目次などをスキップして本文が始まるページ番号を指定します")
        ctrl.addWidget(self._start_spin)

        self._page_total_label = QLabel("/ - ページ")
        ctrl.addWidget(self._page_total_label)

        ctrl.addStretch()
        root.addLayout(ctrl)

        # ── Tree group ────────────────────────────────────────────────────────
        tree_group = QGroupBox("検出されたタイトル・図表")
        tree_layout = QVBoxLayout(tree_group)

        self._tree = QTreeWidget()
        self._tree.setColumnCount(3)
        self._tree.setHeaderLabels(["種別", "ページ", "テキスト"])
        self._tree.setColumnWidth(0, 55)
        self._tree.setColumnWidth(1, 60)
        self._tree.header().setStretchLastSection(True)
        self._tree.setSelectionMode(QTreeWidget.SelectionMode.ExtendedSelection)
        self._tree.itemDoubleClicked.connect(lambda item, col: self._edit_selected())
        tree_layout.addWidget(self._tree)

        root.addWidget(tree_group, stretch=1)

        # ── Edit / Delete buttons ─────────────────────────────────────────────
        btn_row = QHBoxLayout()
        del_btn = QPushButton("削除")
        del_btn.clicked.connect(self._delete_selected)
        edit_btn = QPushButton("編集")
        edit_btn.clicked.connect(self._edit_selected)
        btn_row.addWidget(del_btn)
        btn_row.addWidget(edit_btn)
        btn_row.addStretch()
        root.addLayout(btn_row)

        # ── Status / Save ─────────────────────────────────────────────────────
        bottom = QHBoxLayout()
        self._status_label = QLabel("入力 PDF を選択してください")
        save_btn = QPushButton("しおりを追加して保存")
        save_btn.setObjectName("primary")
        save_btn.clicked.connect(self._process)
        bottom.addWidget(self._status_label, stretch=1)
        bottom.addWidget(save_btn)
        root.addLayout(bottom)

        QShortcut(QKeySequence(Qt.Key.Key_Delete), self._tree).activated.connect(
            self._delete_selected
        )

    # ── File dialogs ──────────────────────────────────────────────────────────

    def _pick_input(self):
        dlg = QFileDialog(self, "入力 PDF を選択")
        dlg.setAcceptMode(QFileDialog.AcceptMode.AcceptOpen)
        dlg.setNameFilter("PDF ファイル (*.pdf);;すべてのファイル (*)")
        dlg.setOption(QFileDialog.Option.DontUseNativeDialog, True)
        dlg.setLabelText(QFileDialog.DialogLabel.Accept, "Open")
        if dlg.exec() != QFileDialog.DialogCode.Accepted:
            return
        files = dlg.selectedFiles()
        if not files:
            return
        p = files[0]
        self._in_edit.setText(p)
        if not self._out_edit.text():
            self._out_edit.setText(default_output(p))
        self._update_page_info(p)

    def _pick_output(self):
        dlg = QFileDialog(self, "出力 PDF の保存先")
        dlg.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
        dlg.setNameFilter("PDF ファイル (*.pdf);;すべてのファイル (*)")
        dlg.setDefaultSuffix("pdf")
        dlg.setOption(QFileDialog.Option.DontUseNativeDialog, True)
        dlg.setOption(QFileDialog.Option.DontConfirmOverwrite, True)
        dlg.setLabelText(QFileDialog.DialogLabel.Accept, "Open")
        if dlg.exec() == QFileDialog.DialogCode.Accepted:
            files = dlg.selectedFiles()
            if files:
                self._out_edit.setText(files[0])

    # ── Detection ─────────────────────────────────────────────────────────────

    def _update_page_info(self, path: str) -> None:
        try:
            doc = fitz.open(path)
            n = len(doc)
            doc.close()
        except Exception:
            return
        self._start_spin.setMaximum(n)
        self._start_spin.setValue(1)
        self._page_total_label.setText(f"/ {n} ページ")

    def _detect(self):
        path = self._in_edit.text().strip()
        if not path:
            QMessageBox.warning(self, "未選択", "入力 PDF を指定してください。")
            return
        if not os.path.isfile(path):
            QMessageBox.critical(self, "エラー", f"ファイルが見つかりません:\n{path}")
            return

        start_page = self._start_spin.value() - 1   # 1-indexed UI → 0-indexed

        try:
            doc = fitz.open(path)
            entries = extract_titles(doc, start_page=start_page)
            doc.close()
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"PDF 読み込みエラー:\n{e}")
            return

        self._entries = promote_levels(entries)
        self._refresh_tree()

        headings = sum(1 for e in self._entries if e.kind == "heading")
        figures  = sum(1 for e in self._entries if e.kind == "figure")
        tables   = sum(1 for e in self._entries if e.kind == "table")
        parts = [f"見出し {headings} 件"]
        if figures: parts.append(f"図 {figures} 件")
        if tables:  parts.append(f"表 {tables} 件")
        skip = f"（{start_page} ページをスキップ）" if start_page > 0 else ""
        self._status_label.setText("、".join(parts) + "を検出しました" + skip)

    # ── Tree helpers ──────────────────────────────────────────────────────────

    def _refresh_tree(self):
        self._tree.clear()

        headings  = [(i, e) for i, e in enumerate(self._entries) if e.kind == "heading"]
        figtables = [(i, e) for i, e in enumerate(self._entries) if e.kind in ("figure", "table")]

        for idx, e in headings:
            indent = "　" * (e.disp_level - 1)
            item = QTreeWidgetItem([f"H{e.disp_level}", str(e.page + 1), indent + e.text])
            item.setTextAlignment(0, Qt.AlignmentFlag.AlignCenter)
            item.setTextAlignment(1, Qt.AlignmentFlag.AlignCenter)
            item.setData(0, Qt.ItemDataRole.UserRole, idx)
            self._tree.addTopLevelItem(item)

        if figtables:
            sep = QTreeWidgetItem(["", "", "── 図表のしおり ──"])
            sep.setFlags(Qt.ItemFlag.ItemIsEnabled)   # visible but not selectable
            for col in range(3):
                sep.setForeground(col, _COL_SEP_FG)
                sep.setBackground(col, _COL_SEP_BG)
            sep.setData(0, Qt.ItemDataRole.UserRole, -1)
            self._tree.addTopLevelItem(sep)

            for idx, e in figtables:
                label = "図" if e.kind == "figure" else "表"
                color = _COL_FIGURE if e.kind == "figure" else _COL_TABLE
                item = QTreeWidgetItem([label, str(e.page + 1), e.text])
                item.setTextAlignment(0, Qt.AlignmentFlag.AlignCenter)
                item.setTextAlignment(1, Qt.AlignmentFlag.AlignCenter)
                item.setForeground(0, color)
                item.setForeground(2, color)
                item.setData(0, Qt.ItemDataRole.UserRole, idx)
                self._tree.addTopLevelItem(item)

    def _selected_indices(self) -> List[int]:
        indices = []
        for item in self._tree.selectedItems():
            idx = item.data(0, Qt.ItemDataRole.UserRole)
            if idx is not None and idx >= 0:
                indices.append(idx)
        return indices

    def _delete_selected(self):
        indices = sorted(self._selected_indices(), reverse=True)
        if not indices:
            return
        for i in indices:
            del self._entries[i]
        self._entries = promote_levels(self._entries)
        self._refresh_tree()
        h = sum(1 for e in self._entries if e.kind == "heading")
        f = sum(1 for e in self._entries if e.kind in ("figure", "table"))
        self._status_label.setText(f"見出し {h} 件、図表 {f} 件が登録されています")

    def _edit_selected(self):
        indices = self._selected_indices()
        if not indices:
            return
        if len(indices) > 1:
            QMessageBox.information(self, "情報", "編集は 1 件ずつ選択してください。")
            return
        idx = indices[0]
        entry = self._entries[idx]
        dlg = EditDialog(entry, self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            t = dlg.get_text()
            if t:
                self._entries[idx].text = t
                self._refresh_tree()

    # ── Process ───────────────────────────────────────────────────────────────

    def _process(self):
        in_p  = self._in_edit.text().strip()
        out_p = self._out_edit.text().strip()

        if not in_p:
            QMessageBox.warning(self, "未選択", "入力 PDF を指定してください。")
            return
        if not os.path.isfile(in_p):
            QMessageBox.critical(self, "エラー", f"ファイルが見つかりません:\n{in_p}")
            return
        if not out_p:
            QMessageBox.warning(self, "未選択", "出力 PDF を指定してください。")
            return
        if not self._entries:
            QMessageBox.warning(
                self, "しおりなし",
                "追加するしおりがありません。\n先に「タイトルを検出」を実行してください。"
            )
            return

        if os.path.abspath(in_p) == os.path.abspath(out_p):
            out_p = default_output(in_p)
            self._out_edit.setText(out_p)

        try:
            n = save_pdf_with_bookmarks(in_p, out_p, self._entries)
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"保存エラー:\n{e}")
            return

        intended = len(build_toc(self._entries))
        if n == 0:
            QMessageBox.warning(
                self, "警告",
                f"ファイルは保存されましたが、しおりが書き込まれませんでした。\n"
                f"（試行 {intended} 件 → 確認 0 件）\n\n{out_p}"
            )
        elif n < intended:
            QMessageBox.warning(
                self, "部分成功",
                f"{n} 件のしおりを追加しました（試行 {intended} 件中）。\n\n{out_p}\n\n"
                "※ PDFビューア左パネルでしおりをご確認ください。"
            )
        else:
            QMessageBox.information(
                self, "完了",
                f"{n} 件のしおりを追加しました。\n\n{out_p}\n\n"
                "※ PDFビューア左パネルでしおりをご確認ください。"
            )
        self._status_label.setText(f"{n} 件のしおりを追加しました")


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
