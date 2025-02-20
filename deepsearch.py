#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Deepsearch - 개선된 UI/UX (Design System 적용)
progress_var(Progressbar)가 우측하단에 고정되도록 수정
"""

import os
import re
import time
import logging
import subprocess
import threading
import zipfile
import xml.etree.ElementTree as ET
from abc import ABC, abstractmethod
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Optional, Callable, Tuple
from datetime import datetime

import fitz  # PyMuPDF
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# -------------------------------------------------------------
# 로깅 설정
# -------------------------------------------------------------
logging.basicConfig(level=logging.INFO, format='[%(levelname)s] %(message)s')
logger = logging.getLogger(__name__)

# -------------------------------------------------------------
# [패치] hwp5 모듈의 FILETIME 관련 OverflowError 보완
# -------------------------------------------------------------
try:
    import hwp5.msoleprops as msoleprops
    from datetime import datetime

    original_property_str = msoleprops.Property.__str__

    def patched_property_str(self):
        try:
            return original_property_str(self)
        except OverflowError:
            return ""
    msoleprops.Property.__str__ = patched_property_str

    if hasattr(msoleprops.Property, "datetime"):
        original_datetime_getter = msoleprops.Property.datetime.fget

        def patched_datetime(self):
            try:
                return original_datetime_getter(self)
            except OverflowError:
                return None
        msoleprops.Property.datetime = property(patched_datetime)
    logger.info("hwp5 모듈 패치 완료.")
except Exception as e:
    logger.error("hwp5 패치 실패: " + str(e))

# -------------------------------------------------------------
# 드래그 앤 드롭 지원 (tkinterdnd2 사용, 없으면 기본 Tk)
# -------------------------------------------------------------
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DnDTk = TkinterDnD.Tk
except ImportError:
    DnDTk = tk.Tk

# -------------------------------------------------------------
# 디자인 시스템: COLORS, FONTS
# -------------------------------------------------------------
COLORS = {
    'primary': '#4A90E2',
    'accent': '#357ABD',
    'background': '#FFFFFF',
    'text_primary': '#333333',
    'text_secondary': '#777777',
    'border_light': '#E0E0E0',
    'hover_background': '#F5F5F5'
}

FONTS = {
    'default': 'Segoe UI',
    'sizes': {
        'small': 10,
        'normal': 12,
        'large': 14,
        'heading': 16
    },
    'weights': {
        'regular': 'normal',
        'bold': 'bold'
    }
}

def create_themed_style(root: tk.Tk) -> None:
    style = ttk.Style(root)
    style.theme_use('default')

    style.configure('Card.TFrame',
        background=COLORS['background'],
        relief='flat',
        borderwidth=1
    )

    style.configure('TButton',
        font=(FONTS['default'], FONTS['sizes']['normal'], FONTS['weights']['bold']),
        background=COLORS['primary'],
        foreground='#FFFFFF',
        padding=(10, 5),
        borderwidth=0,
        relief='flat'
    )
    style.map('TButton',
        background=[
            ('active', COLORS['accent']),
            ('disabled', '#B0B0B0')
        ],
        foreground=[('disabled', 'white')]
    )

    style.configure('Search.TButton',
        font=(FONTS['default'], FONTS['sizes']['normal'], FONTS['weights']['bold']),
        background=COLORS['primary'],
        foreground='#FFFFFF',
        padding=(15, 8),
        borderwidth=0
    )
    style.map('Search.TButton',
        background=[
            ('active', COLORS['accent']),
            ('disabled', '#B0B0B0')
        ],
        foreground=[('disabled', 'white')]
    )

    style.configure('TEntry',
        font=(FONTS['default'], FONTS['sizes']['normal']),
        padding=(5, 5)
    )

    style.configure('Filter.TCheckbutton',
        font=(FONTS['default'], FONTS['sizes']['normal']),
        background=COLORS['background']
    )

    style.configure('TLabel',
        background=COLORS['background'],
        foreground=COLORS['text_primary'],
        font=(FONTS['default'], FONTS['sizes']['normal'])
    )

# -------------------------------------------------------------
# 파일 파서 추상 클래스 및 전용 파서
# -------------------------------------------------------------
class BaseFileParser(ABC):
    @abstractmethod
    def parse(self, file_path: str) -> str:
        pass

class HWPParser(BaseFileParser):
    def parse(self, file_path: str) -> str:
        try:
            output = subprocess.check_output(
                ["hwp5txt", file_path],
                encoding="utf-8",
                errors="ignore",
                stderr=subprocess.STDOUT
            )
            return output.strip()
        except subprocess.CalledProcessError as e:
            logger.error(f"[HWPParser] 파싱 오류 ({file_path}): {e.output.strip()}")
            return ""
        except Exception as e:
            logger.error(f"[HWPParser] 예외 발생 ({file_path}): {e}")
            return ""

class HWPXParser(BaseFileParser):
    def parse(self, file_path: str) -> str:
        texts = []
        try:
            with zipfile.ZipFile(file_path, 'r') as z:
                for name in z.namelist():
                    if name.startswith("Contents/") and name.lower().endswith(".xml"):
                        try:
                            with z.open(name) as f:
                                tree = ET.parse(f)
                                root = tree.getroot()
                                for elem in root.iter():
                                    if '}' in elem.tag:
                                        elem.tag = elem.tag.split('}', 1)[1]
                                text = "".join(root.itertext())
                                texts.append(text)
                        except Exception as xe:
                            logger.error(f"[HWPXParser] XML 파싱 오류 ({name} in {file_path}): {xe}")
            return "\n".join(texts).strip()
        except zipfile.BadZipFile:
            logger.error(f"[HWPXParser] 파일 형식 오류 (올바른 hwpx 파일이 아님): {file_path}")
            return ""
        except Exception as e:
            logger.error(f"[HWPXParser] 파일 파싱 오류 ({file_path}): {e}")
            return ""

class PDFParser(BaseFileParser):
    def parse(self, file_path: str) -> str:
        text_content = []
        try:
            with fitz.open(file_path) as pdf:
                for page in pdf:
                    text_content.append(page.get_text())
        except Exception as e:
            logger.error(f"[PDFParser] 파일 파싱 오류 ({file_path}): {e}")
            return ""
        return "\n".join(text_content)

class ExcelParser(BaseFileParser):
    def parse(self, file_path: str) -> str:
        text_content = []
        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                df = xls.parse(sheet_name=sheet_name, header=None)
                df = df.fillna("").astype(str)
                text_content.append(f"[Sheet: {sheet_name}]")
                for row in df.values.tolist():
                    text_content.append(" ".join(row))
        except Exception as e:
            logger.error(f"[ExcelParser] 파일 파싱 오류 ({file_path}): {e}")
            return ""
        return "\n".join(text_content)

class ParserFactory:
    @staticmethod
    def get_parser(file_path: str) -> BaseFileParser:
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".hwp":
            return HWPParser()
        elif ext == ".hwpx":
            return HWPXParser()
        elif ext == ".pdf":
            return PDFParser()
        elif ext in (".xls", ".xlsx"):
            return ExcelParser()
        else:
            raise ValueError(f"지원하지 않는 파일 확장자: {ext}")

class FileParser:
    @staticmethod
    def parse_file(file_path: str) -> str:
        try:
            parser = ParserFactory.get_parser(file_path)
            return parser.parse(file_path)
        except Exception as e:
            logger.error(f"[FileParser] 파서 선택 오류 ({file_path}): {e}")
            return ""

# -------------------------------------------------------------
# IndexManager: Whoosh 인덱스 관리
# -------------------------------------------------------------
from whoosh.index import create_in, open_dir
from whoosh.fields import Schema, ID, TEXT, DATETIME
from whoosh.qparser import MultifieldParser, OrGroup, AndGroup
from whoosh.analysis import RegexAnalyzer

class IndexManager:
    def __init__(self, index_dir: str = "indexdir") -> None:
        self.index_dir = index_dir
        self.schema = Schema(
            path=ID(stored=True, unique=True),
            filename=TEXT(stored=True),
            extension=TEXT(stored=True),
            content=TEXT(analyzer=RegexAnalyzer(), stored=True),
            modified=DATETIME(stored=True)
        )
        if not os.path.exists(self.index_dir):
            os.makedirs(self.index_dir)
            create_in(self.index_dir, self.schema)
        self.ix = open_dir(self.index_dir)

    def clear_index(self) -> None:
        try:
            self.ix.close()
        except Exception:
            pass
        for f in os.listdir(self.index_dir):
            filepath = os.path.join(self.index_dir, f)
            try:
                os.remove(filepath)
            except Exception as e:
                logger.warning(f"Could not remove {filepath}: {e}")
                time.sleep(0.5)
                try:
                    os.remove(filepath)
                except Exception as e2:
                    logger.warning(f"Retry failed for {filepath}: {e2}")
        create_in(self.index_dir, self.schema)
        self.ix = open_dir(self.index_dir)

    def index_files(self,
                    file_paths: List[str],
                    progress_callback: Optional[Callable[[int, int, float], None]] = None,
                    cancel_callback: Optional[Callable[[], bool]] = None,
                    max_workers: int = 6) -> None:
        def parse_job(fpath: str) -> Tuple[str, Optional[str], Optional[str], Optional[str], Optional[datetime], Optional[Exception]]:
            try:
                extension = os.path.splitext(fpath)[1].lower()
                filename = os.path.basename(fpath)
                content = FileParser.parse_file(fpath)
                modified = datetime.fromtimestamp(os.path.getmtime(fpath))
                return (fpath, extension, filename, content, modified, None)
            except Exception as e:
                return (fpath, None, None, None, None, e)

        results = []
        total_files = len(file_paths)
        start_parse_time = time.time()

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_path = {executor.submit(parse_job, f): f for f in file_paths}
            for i, future in enumerate(as_completed(future_to_path), start=1):
                if cancel_callback and cancel_callback():
                    logger.info("색인 중단 요청됨.")
                    break

                fpath = future_to_path[future]
                try:
                    fpath, extension, filename, content, modified, err = future.result()
                    if err:
                        logger.error(f"[index_files] 파싱 오류: {fpath}, {err}")
                        if progress_callback:
                            progress_callback(i, total_files, 0)
                        continue
                    results.append((fpath, extension, filename, content, modified))
                except Exception as exc:
                    logger.error(f"[index_files] 예외 발생: {fpath}, {exc}")
                finally:
                    if progress_callback:
                        elapsed = time.time() - start_parse_time
                        progress_callback(i, total_files, elapsed)

        end_parse_time = time.time()
        parse_duration = end_parse_time - start_parse_time

        start_index_time = time.time()
        writer = self.ix.writer()
        for (fpath, extension, filename, content, modified) in results:
            writer.update_document(
                path=fpath,
                filename=filename,
                extension=extension,
                content=content,
                modified=modified
            )
        writer.commit()
        end_index_time = time.time()
        index_duration = end_index_time - start_index_time
        total_duration = parse_duration + index_duration
        logger.info(f"멀티스레드 파싱: {parse_duration:.2f}초, 인덱싱: {index_duration:.2f}초, 총: {total_duration:.2f}초")

    def search(self, query_str: str, and_mode: bool = False, sort_by: str = "relevance") -> List[dict]:
        with self.ix.searcher() as searcher:
            if and_mode:
                parser = MultifieldParser(["filename", "content"], schema=self.ix.schema, group=AndGroup)
            else:
                parser = MultifieldParser(["filename", "content"], schema=self.ix.schema, group=OrGroup)
            query = parser.parse(query_str)

            if sort_by == "date":
                results = searcher.search(query, limit=50, sortedby="modified", reverse=True)
            else:
                results = searcher.search(query, limit=50)

            results.fragmenter.charlimit = None
            hits = []
            for r in results:
                hits.append({
                    "path": r["path"],
                    "filename": r["filename"],
                    "extension": r["extension"],
                    "content": r["content"],
                    "modified": r["modified"]
                })
            return hits

# -------------------------------------------------------------
# OptimizedApp: Tkinter GUI (Progressbar 위치 고정)
# -------------------------------------------------------------
class OptimizedApp(DnDTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Deepsearch")
        self.geometry("1200x800")
        self.configure(bg=COLORS['background'])

        create_themed_style(self)

        self.monitor_dirs: List[str] = []
        self.idx_manager = IndexManager()
        self.current_results: List[dict] = []
        self.is_indexing = False
        self.index_cancelled = False
        self.last_index_time: Optional[datetime] = None
        self.auto_index_interval = 0
        self.current_context_path: Optional[str] = None

        self._setup_ui()
        self._bind_events()
        self.after(100, self._show_initial_guide)
        self.update_idletasks()
        self._update_shortcut_info()

    def _setup_ui(self) -> None:
        # ----------------------------
        # 상단 검색 영역
        # ----------------------------
        self.search_frame = ttk.Frame(self, style="Card.TFrame")
        self.search_frame.pack(fill=tk.X, padx=20, pady=15)
        self.search_frame.configure(padding=10)

        search_container = ttk.Frame(self.search_frame, style="Card.TFrame")
        search_container.pack(fill=tk.X, padx=5, pady=5)

        search_icon = ttk.Label(
            search_container,
            text="🔍",
            font=(FONTS['default'], FONTS['sizes']['large'])
        )
        search_icon.pack(side=tk.LEFT, padx=(10, 5), pady=10)

        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(
            search_container, textvariable=self.search_var,
            style="TEntry",
            font=(FONTS['default'], FONTS['sizes']['large'])
        )
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10, pady=10)

        # placeholder
        if not self.monitor_dirs:
            self.search_entry.insert(0, "먼저 검색폴더를 추가해주세요")
            self.search_entry.config(foreground="gray")
        else:
            self.search_entry.insert(0, "검색어를 입력하세요")
            self.search_entry.config(foreground="gray")

        self.search_button = ttk.Button(search_container, text="검색", style="Search.TButton", command=self.on_search)
        self.search_button.pack(side=tk.LEFT, padx=10, pady=10)

        options_container = ttk.Frame(self.search_frame, style="Card.TFrame")
        options_container.pack(fill=tk.X, padx=5, pady=(0,10))

        ttk.Label(options_container, text="정렬:").pack(side=tk.LEFT, padx=(10,2))
        self.sort_var = tk.StringVar(value="관련도순")
        self.sort_combobox = ttk.Combobox(
            options_container, textvariable=self.sort_var, width=10,
            values=["관련도순", "날짜순"], state="readonly"
        )
        self.sort_combobox.pack(side=tk.LEFT, padx=5)

        ttk.Label(options_container, text="검색 모드:").pack(side=tk.LEFT, padx=(20,2))
        self.mode_var = tk.StringVar(value="OR")
        self.mode_combobox = ttk.Combobox(
            options_container, textvariable=self.mode_var, width=6,
            values=["OR", "AND"], state="readonly"
        )
        self.mode_combobox.pack(side=tk.LEFT, padx=5)

        ttk.Label(options_container, text="파일 유형:").pack(side=tk.LEFT, padx=(20,2))
        self.filter_hwp = tk.BooleanVar(value=True)
        self.filter_hwpx = tk.BooleanVar(value=True)
        self.filter_pdf = tk.BooleanVar(value=True)
        self.filter_excel = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_container, text="HWP", variable=self.filter_hwp, style="Filter.TCheckbutton").pack(side=tk.LEFT, padx=2)
        ttk.Checkbutton(options_container, text="HWPX", variable=self.filter_hwpx, style="Filter.TCheckbutton").pack(side=tk.LEFT, padx=2)
        ttk.Checkbutton(options_container, text="PDF", variable=self.filter_pdf, style="Filter.TCheckbutton").pack(side=tk.LEFT, padx=2)
        ttk.Checkbutton(options_container, text="Excel", variable=self.filter_excel, style="Filter.TCheckbutton").pack(side=tk.LEFT, padx=2)

        toolbar_frame = ttk.Frame(search_container, style="Card.TFrame")
        toolbar_frame.pack(side=tk.RIGHT, padx=10)

        self.reindex_button = ttk.Button(toolbar_frame, text="🔄 색인", command=self.reindex_files)
        self.reindex_button.pack(side=tk.LEFT, padx=5)

        self.settings_button = ttk.Button(toolbar_frame, text="⚙️ 검색폴더추가", command=self.open_settings)
        self.settings_button.pack(side=tk.LEFT, padx=5)

        # ----------------------------
        # 결과 영역 (스크롤 가능)
        # ----------------------------
        results_frame = ttk.Frame(self, style="Card.TFrame")
        results_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0,15))

        self.results_canvas = tk.Canvas(results_frame, bg=COLORS['background'], highlightthickness=0)
        self.results_scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=self.results_canvas.yview)
        self.results_canvas.configure(yscrollcommand=self.results_scrollbar.set)
        self.results_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.results_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.results_container = ttk.Frame(self.results_canvas, style="Card.TFrame")
        self.results_canvas.create_window((0,0), window=self.results_container, anchor="nw")

        self.results_container.bind("<Configure>", lambda e: self.results_canvas.configure(scrollregion=self.results_canvas.bbox("all")))

        # ----------------------------
        # 하단 상태바
        # ----------------------------
        status_frame = ttk.Frame(self, style="Card.TFrame")
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)

        # [변경] 진행바를 오른쪽(우측 하단)에 고정하기 위해 pack을 분리
        # 상단(왼쪽) 상태 영역
        left_status_frame = ttk.Frame(status_frame, style="Card.TFrame")
        left_status_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.status_var = tk.StringVar(value="준비")
        status_label = ttk.Label(left_status_frame, textvariable=self.status_var)
        status_label.pack(side=tk.LEFT, padx=20, pady=2)

        # 우측(하단) 진행상황 영역
        right_status_frame = ttk.Frame(status_frame, style="Card.TFrame")
        right_status_frame.pack(side=tk.RIGHT)

        self.progress_var = tk.IntVar()
        self.progress_bar = ttk.Progressbar(right_status_frame, mode="determinate",
                                            variable=self.progress_var, length=200)
        self.progress_bar.pack(side=tk.RIGHT, padx=20, pady=2)

        # 마지막 색인, 단축키 안내, 개발자 라벨 등은 왼쪽 영역 아래쪽에 배치
        bottom_status_subframe = ttk.Frame(left_status_frame, style="Card.TFrame")
        bottom_status_subframe.pack(side=tk.BOTTOM, fill=tk.X)

        self.last_index_label = ttk.Label(bottom_status_subframe, text="마지막 색인: 없음", foreground=COLORS['text_secondary'])
        self.last_index_label.pack(side=tk.LEFT, padx=20, pady=2)

        self.shortcut_label = ttk.Label(bottom_status_subframe, text="", font=(FONTS['default'], FONTS['sizes']['small']))
        self.shortcut_label.pack(side=tk.LEFT, padx=20, pady=2)

        developer_label = ttk.Label(
            bottom_status_subframe,
            text="developed by 김홍교",
            font=(FONTS['default'], FONTS['sizes']['small'], FONTS['weights']['bold']),
            foreground=COLORS['text_secondary']
        )
        developer_label.pack(side=tk.RIGHT, padx=20, pady=2)

    def _bind_events(self) -> None:
        self.search_entry.bind("<Return>", self.on_search)
        self.search_entry.bind("<FocusIn>", self._clear_placeholder)
        self.search_entry.bind("<FocusOut>", self._add_placeholder)
        self.bind("<Control-f>", self._focus_search_entry)
        self.bind("<Escape>", self._clear_search_entry)

        self.results_canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.results_container.bind("<MouseWheel>", self._on_mousewheel)
        self.results_canvas.bind("<Button-4>", self._on_mousewheel)
        self.results_canvas.bind("<Button-5>", self._on_mousewheel)
        self.results_container.bind("<Button-4>", self._on_mousewheel)
        self.results_container.bind("<Button-5>", self._on_mousewheel)

    def _update_shortcut_info(self) -> None:
        info = "단축키: Ctrl+F (검색창 포커스), Esc (검색창 초기화)"
        self.shortcut_label.config(text=info)

    # Placeholder 제어
    def _clear_placeholder(self, event: tk.Event) -> None:
        current = self.search_entry.get()
        if current in ["검색어를 입력하세요", "먼저 검색폴더를 추가해주세요"]:
            self.search_entry.delete(0, tk.END)
            self.search_entry.config(foreground=COLORS['text_primary'])

    def _add_placeholder(self, event: tk.Event) -> None:
        if not self.search_entry.get():
            if not self.monitor_dirs:
                self.search_entry.insert(0, "먼저 검색폴더를 추가해주세요")
            else:
                self.search_entry.insert(0, "검색어를 입력하세요")
            self.search_entry.config(foreground="gray")

    def _focus_search_entry(self, event: tk.Event = None) -> None:
        self.search_entry.focus_set()
        self.search_entry.selection_range(0, tk.END)

    def _clear_search_entry(self, event: tk.Event = None) -> None:
        self.search_entry.delete(0, tk.END)

    # 마우스휠 스크롤
    def _on_mousewheel(self, event: tk.Event) -> None:
        if event.num == 4:
            self.results_canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.results_canvas.yview_scroll(1, "units")
        else:
            scroll_steps = int(-event.delta / 120)
            self.results_canvas.yview_scroll(scroll_steps, "units")

    def _show_initial_guide(self) -> None:
        if self.results_container.winfo_children():
            return
        for child in self.results_container.winfo_children():
            child.destroy()

        guide_frame = ttk.Frame(self.results_container, style="Card.TFrame", padding=20)
        guide_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        guide_title = ttk.Label(
            guide_frame,
            text="시작하기",
            font=(FONTS['default'], FONTS['sizes']['heading'], FONTS['weights']['bold'])
        )
        guide_title.pack(pady=10)

        guide_text = (
            "👉 1. '⚙️ 검색폴더추가' 버튼을 눌러 검색할 폴더를 추가하세요.\n"
            "👉 2. '🔄 색인' 버튼을 눌러 파일 색인을 진행하세요.\n"
            "👉 3. 검색창에 키워드를 입력하여 검색하세요.\n"
            "단축키: Ctrl+F (검색창 포커스), Esc (검색창 초기화)"
        )
        guide_label = ttk.Label(
            guide_frame,
            text=guide_text,
            font=(FONTS['default'], FONTS['sizes']['large']),
            wraplength=800,
            justify="left"
        )
        guide_label.pack(pady=10)

    # --------------------------------------------------
    # 검색 및 결과 표시
    # --------------------------------------------------
    def on_search(self, event: tk.Event = None) -> None:
        query = self.search_var.get().strip()
        if not query or query in ["검색어를 입력하세요", "먼저 검색폴더를 추가해주세요"]:
            self._show_initial_guide()
            return

        and_mode = (self.mode_var.get() == "AND")
        sort_by = "date" if self.sort_var.get() == "날짜순" else "relevance"

        results = self.idx_manager.search(query, and_mode, sort_by)

        filtered = []
        for r in results:
            ext = r["extension"]
            if ext == ".hwp" and not self.filter_hwp.get():
                continue
            if ext == ".hwpx" and not self.filter_hwpx.get():
                continue
            if ext == ".pdf" and not self.filter_pdf.get():
                continue
            if ext in (".xls", ".xlsx") and not self.filter_excel.get():
                continue
            filtered.append(r)

        self.current_results = filtered
        self._update_result_list(query, filtered)

    def _update_result_list(self, query: str, results: List[dict]) -> None:
        for child in self.results_container.winfo_children():
            child.destroy()

        if not results:
            self._show_initial_guide()
            self.status_var.set("검색 결과가 없습니다.")
            return

        for r in results:
            card = ttk.Frame(self.results_container, style="Card.TFrame", padding=10)
            card.pack(fill=tk.X, padx=10, pady=5)

            icon_label = ttk.Label(
                card,
                text=self._get_icon(r["extension"]),
                font=(FONTS['default'], FONTS['sizes']['heading'])
            )
            icon_label.grid(row=0, column=0, rowspan=2, padx=10, sticky="n")

            fname = ttk.Label(
                card,
                text=f"[{r['filename']}]",
                font=(FONTS['default'], FONTS['sizes']['large'], FONTS['weights']['bold'])
            )
            fname.grid(row=0, column=1, sticky="w")
            fname.config(cursor="hand2")

            mod_time = r.get("modified")
            mod_str = mod_time.strftime("%Y-%m-%d %H:%M:%S") if isinstance(mod_time, datetime) else ""
            mod_label = ttk.Label(
                card,
                text=f"수정: {mod_str}",
                font=(FONTS['default'], FONTS['sizes']['small']),
                foreground=COLORS['text_secondary']
            )
            mod_label.grid(row=0, column=2, sticky="e", padx=10)

            snippet = self._generate_snippet(r["content"], query)
            snippet_label = ttk.Label(
                card,
                text=snippet,
                font=(FONTS['default'], FONTS['sizes']['normal']),
                wraplength=800,
                justify="left"
            )
            snippet_label.grid(row=1, column=1, columnspan=2, sticky="w", pady=5)

            # 더블클릭으로 파일 열기
            card.bind("<Double-Button-1>", lambda e, path=r["path"]: self._open_file(path))
            fname.bind("<Double-Button-1>", lambda e, path=r["path"]: self._open_file(path))
            for widget in card.winfo_children():
                widget.bind("<Double-Button-1>", lambda e, path=r["path"]: self._open_file(path))

        self.status_var.set(f"총 {len(results)}개 문서 검색됨")

    def _get_icon(self, extension: str) -> str:
        if extension in (".hwp", ".hwpx"):
            return "📄"
        elif extension == ".pdf":
            return "📑"
        elif extension in (".xls", ".xlsx"):
            return "📊"
        return "🗂️"

    def _generate_snippet(self, content: str, query: str) -> str:
        lines = content.split("\n")
        query_terms = re.findall(r"\b\w+\b", query.lower())
        snippet_line = None
        for line in lines:
            if any(term in line.lower() for term in query_terms):
                snippet_line = line.strip()
                break
        if snippet_line:
            snippet = snippet_line[:100] + "..." if len(snippet_line) > 100 else snippet_line
            for term in query_terms:
                snippet = re.sub(rf'\b({re.escape(term)})\b', r'**\1**', snippet, flags=re.IGNORECASE)
            return f"→ {snippet}"
        return "→ [내용 미리보기 없음]" if lines else ""

    def _open_file(self, path: str) -> None:
        if os.path.exists(path):
            try:
                if os.name == "nt":
                    os.startfile(path)
                else:
                    subprocess.Popen(["xdg-open", path])
            except Exception as e:
                messagebox.showerror("오류", f"파일 열기 실패: {e}")
        else:
            messagebox.showerror("오류", f"파일을 찾을 수 없습니다: {path}")

    # --------------------------------------------------
    # 색인(Re-index)
    # --------------------------------------------------
    def reindex_files(self) -> None:
        if self.is_indexing:
            self.index_cancelled = True
            self.status_var.set("색인 중단 요청 중...")
            return

        if not self.monitor_dirs:
            messagebox.showwarning("경고", "먼저 모니터링 폴더를 설정해주세요.")
            return

        def thread_target() -> None:
            self.is_indexing = True
            self.index_cancelled = False
            self.progress_var.set(0)
            self.status_var.set("색인 초기화 중...")
            self.reindex_button.config(text="⏹ 색인 중단")
            self.settings_button.config(state=tk.DISABLED)
            self.search_button.config(state=tk.DISABLED)
            self.update_idletasks()
            time.sleep(0.3)

            self.idx_manager.clear_index()
            files = self._collect_files()
            total = len(files)
            if total == 0:
                self.status_var.set("색인할 파일이 없습니다.")
                self.is_indexing = False
                self.reindex_button.config(text="🔄 색인")
                self.settings_button.config(state=tk.NORMAL)
                self.search_button.config(state=tk.NORMAL)
                return

            def progress_callback(current: int, total_files: int, elapsed: float) -> None:
                progress = int((current / total_files) * 100)
                remaining = 0
                if current > 0:
                    remaining = int((elapsed / current) * (total_files - current))
                self.progress_var.set(progress)
                filename = os.path.basename(files[current - 1]) if current - 1 < len(files) else ""
                self.status_var.set(f"색인 중... ({current}/{total_files}) {filename} | 남은 시간: {remaining}s")
                self.update_idletasks()

            start_time = time.time()
            self.idx_manager.index_files(
                files,
                progress_callback,
                cancel_callback=lambda: self.index_cancelled,
                max_workers=6
            )
            elapsed_time = time.time() - start_time

            if self.index_cancelled:
                self.status_var.set("색인 중단됨")
            else:
                self.status_var.set(f"색인 완료 ({total}개, {elapsed_time:.2f}초 소요)")
                self.last_index_time = datetime.now()
                self.last_index_label.config(text=f"마지막 색인: {self.last_index_time.strftime('%Y-%m-%d %H:%M:%S')}")

            self.progress_var.set(100)
            self.is_indexing = False
            self.reindex_button.config(text="🔄 색인")
            self.settings_button.config(state=tk.NORMAL)
            self.search_button.config(state=tk.NORMAL)
            self.update_idletasks()

        threading.Thread(target=thread_target, daemon=True).start()

    # --------------------------------------------------
    # 설정(검색폴더) 관련
    # --------------------------------------------------
    def open_settings(self) -> None:
        win = tk.Toplevel(self)
        win.title("모니터링 폴더 설정")
        win.configure(bg=COLORS['background'])
        win.geometry("600x500")
        win.minsize(400, 300)

        list_frame = ttk.Frame(win, style="Card.TFrame")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        ttk.Label(
            list_frame, text="모니터링 폴더 목록",
            font=(FONTS['default'], FONTS['sizes']['normal'], FONTS['weights']['bold'])
        ).pack(pady=(0, 10), anchor=tk.W)

        self.dir_list = tk.Listbox(
            list_frame, selectmode=tk.SINGLE,
            font=(FONTS['default'], FONTS['sizes']['normal']),
            borderwidth=2, relief="flat",
            highlightcolor=COLORS['accent'], highlightthickness=1
        )
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.dir_list.yview)
        hsb = ttk.Scrollbar(list_frame, orient="horizontal", command=self.dir_list.xview)
        self.dir_list.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.dir_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for d in self.monitor_dirs:
            self.dir_list.insert(tk.END, d)

        self.dir_list.bind("<Double-Button-1>", self._remove_folder_event)

        auto_frame = ttk.Frame(win, style="Card.TFrame")
        auto_frame.pack(fill=tk.X, padx=15, pady=10)

        ttk.Label(
            auto_frame, text="자동 색인 주기 (분, 0=미사용):"
        ).pack(side=tk.LEFT)
        self.auto_index_var = tk.StringVar(value=str(self.auto_index_interval))
        auto_spin = ttk.Spinbox(auto_frame, from_=0, to=60, textvariable=self.auto_index_var,
                                width=5)
        auto_spin.pack(side=tk.LEFT, padx=5)

        btn_frame = ttk.Frame(win, style="Card.TFrame")
        btn_frame.pack(fill=tk.X, padx=15, pady=5)
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)

        ttk.Button(btn_frame, text="폴더 추가", command=lambda: self._add_folder(win)).grid(row=0, column=0, sticky=tk.EW, padx=5, pady=8)
        ttk.Button(btn_frame, text="폴더 제거", command=self._remove_folder).grid(row=0, column=1, sticky=tk.EW, padx=5, pady=8)

        info_label = ttk.Label(
            win, text="* 목록에서 폴더를 더블클릭하여 제거할 수도 있습니다.",
            foreground=COLORS['text_secondary']
        )
        info_label.pack(side=tk.BOTTOM, pady=10, anchor=tk.W, padx=15)

        def apply_settings():
            try:
                self.auto_index_interval = int(self.auto_index_var.get())
            except:
                self.auto_index_interval = 0
            messagebox.showinfo("알림", "설정이 적용되었습니다.")
            win.destroy()
            if self.auto_index_interval > 0:
                self._schedule_auto_index()

        ttk.Button(win, text="설정 적용", command=apply_settings).pack(pady=10)

    def _on_drop(self, event: tk.Event) -> None:
        dropped = self.tk.splitlist(event.data)
        for path in dropped:
            if os.path.isdir(path) and path not in self.monitor_dirs:
                self.monitor_dirs.append(path)
                self.dir_list.insert(tk.END, path)

    def _remove_folder_event(self, event: tk.Event) -> None:
        self._remove_folder()

    def _add_folder(self, parent: tk.Toplevel) -> None:
        if len(self.monitor_dirs) >= 5:
            messagebox.showwarning("경고", "최대 5개 폴더까지 추가 가능합니다.")
            return
        folder = filedialog.askdirectory(parent=parent, title="모니터링 폴더 선택")
        if folder and folder not in self.monitor_dirs:
            self.monitor_dirs.append(folder)
            self.dir_list.insert(tk.END, folder)
            if self.monitor_dirs and self.search_entry.get() == "먼저 검색폴더를 추가해주세요":
                self.search_entry.delete(0, tk.END)
                self.search_entry.insert(0, "검색어를 입력하세요")
                self.search_entry.config(foreground=COLORS['text_primary'])

    def _remove_folder(self) -> None:
        sel = self.dir_list.curselection()
        if not sel:
            messagebox.showwarning("경고", "제거할 폴더를 선택해주세요.")
            return
        idx = sel[0]
        removed_folder = self.monitor_dirs.pop(idx)
        self.dir_list.delete(idx)
        messagebox.showinfo("알림", f"폴더 '{removed_folder}'가 제거되었습니다.")
        if not self.monitor_dirs:
            self.search_entry.delete(0, tk.END)
            self.search_entry.insert(0, "먼저 검색폴더를 추가해주세요")
            self.search_entry.config(foreground="gray")

    def _collect_files(self) -> List[str]:
        exts = [".hwp", ".hwpx", ".pdf", ".xls", ".xlsx"]
        files = []
        for d in self.monitor_dirs:
            for root, _, files_in_dir in os.walk(d):
                for f in files_in_dir:
                    if os.path.splitext(f)[1].lower() in exts:
                        files.append(os.path.join(root, f))
        return files

    def _schedule_auto_index(self) -> None:
        interval_ms = self.auto_index_interval * 60 * 1000
        if interval_ms > 0:
            self.after(interval_ms, self._auto_index)

    def _auto_index(self) -> None:
        if not self.is_indexing:
            self.reindex_files()
        self._schedule_auto_index()

    def run(self) -> None:
        self.mainloop()

# 메인
if __name__ == "__main__":
    app = OptimizedApp()
    app.run()
