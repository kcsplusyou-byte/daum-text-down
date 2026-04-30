import queue
import threading
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
import tkinter as tk

try:
    import main as crawler
except Exception as import_error:
    crawler = None
    IMPORT_ERROR = import_error
else:
    IMPORT_ERROR = None


class CafeCollectorGui:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("다음카페 Re글 수집 프로그램")
        self.root.geometry("820x620")
        self.root.minsize(760, 560)

        self.driver = None
        self.worker: threading.Thread | None = None
        self.messages: queue.Queue[tuple[str, object]] = queue.Queue()

        self.board_url_var = tk.StringVar(value=crawler.BOARD_URL)
        self.start_page_var = tk.StringVar()
        self.end_page_var = tk.StringVar()
        self.output_path_var = tk.StringVar(value=str(Path.cwd() / crawler.OUTPUT_FILE))
        self.status_var = tk.StringVar(value="대기 중")

        self._build_ui()
        self._poll_messages()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_ui(self) -> None:
        root_frame = ttk.Frame(self.root, padding=14)
        root_frame.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        root_frame.columnconfigure(1, weight=1)
        root_frame.rowconfigure(6, weight=1)

        title = ttk.Label(root_frame, text="다음카페 Re글 수집 프로그램", font=("맑은 고딕", 15, "bold"))
        title.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 14))

        ttk.Label(root_frame, text="게시판 주소").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(root_frame, textvariable=self.board_url_var).grid(
            row=1, column=1, columnspan=3, sticky="ew", pady=4
        )

        range_frame = ttk.LabelFrame(root_frame, text="수집 페이지 범위", padding=10)
        range_frame.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(8, 8))
        range_frame.columnconfigure(1, weight=1)
        range_frame.columnconfigure(3, weight=1)

        ttk.Label(range_frame, text="시작 페이지").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(range_frame, textvariable=self.start_page_var, width=12).grid(
            row=0, column=1, sticky="w", padx=(0, 20)
        )
        ttk.Label(range_frame, text="종료 페이지").grid(row=0, column=2, sticky="w", padx=(0, 8))
        ttk.Entry(range_frame, textvariable=self.end_page_var, width=12).grid(row=0, column=3, sticky="w")
        ttk.Label(
            range_frame,
            text="비워두면 현재 페이지부터 감지된 마지막 페이지까지 수집합니다.",
            foreground="#555555",
        ).grid(row=1, column=0, columnspan=4, sticky="w", pady=(8, 0))

        ttk.Label(root_frame, text="저장 파일").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(root_frame, textvariable=self.output_path_var).grid(row=3, column=1, sticky="ew", pady=4)
        ttk.Button(root_frame, text="찾기", command=self.browse_output).grid(
            row=3, column=2, sticky="ew", padx=(8, 0), pady=4
        )

        button_frame = ttk.Frame(root_frame)
        button_frame.grid(row=4, column=0, columnspan=4, sticky="ew", pady=(12, 8))
        button_frame.columnconfigure(5, weight=1)

        self.open_button = ttk.Button(button_frame, text="브라우저 열기", command=self.open_browser)
        self.open_button.grid(row=0, column=0, padx=(0, 8))

        self.collect_button = ttk.Button(button_frame, text="수집 시작", command=self.start_collect, state="disabled")
        self.collect_button.grid(row=0, column=1, padx=(0, 8))

        self.close_button = ttk.Button(button_frame, text="브라우저 닫기", command=self.close_browser, state="disabled")
        self.close_button.grid(row=0, column=2, padx=(0, 8))

        ttk.Label(button_frame, textvariable=self.status_var, anchor="e").grid(row=0, column=5, sticky="e")

        guide = (
            "1. 브라우저 열기 -> 2. 열린 Chrome에서 직접 로그인 -> "
            "3. 게시판 목록 확인 -> 4. 페이지 범위 입력 -> 5. 수집 시작"
        )
        ttk.Label(root_frame, text=guide, foreground="#444444").grid(
            row=5, column=0, columnspan=4, sticky="w", pady=(0, 8)
        )

        self.log_text = scrolledtext.ScrolledText(root_frame, height=14, wrap="word", state="disabled")
        self.log_text.grid(row=6, column=0, columnspan=4, sticky="nsew")

    def browse_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="저장할 워드파일 선택",
            initialfile=Path(self.output_path_var.get()).name,
            defaultextension=".docx",
            filetypes=[("Word 문서", "*.docx"), ("모든 파일", "*.*")],
        )
        if path:
            self.output_path_var.set(path)

    def open_browser(self) -> None:
        if self.driver is not None:
            messagebox.showinfo("안내", "브라우저가 이미 열려 있습니다.")
            return

        board_url = self.board_url_var.get().strip()
        if not board_url:
            messagebox.showerror("입력 오류", "게시판 주소를 입력하세요.")
            return

        self._run_worker(lambda: self._open_browser_task(board_url), "브라우저 실행 중")

    def _open_browser_task(self, board_url: str) -> None:
        crawler.BOARD_URL = board_url
        crawler.log = self._thread_log
        self._thread_log("Chrome 브라우저를 실행합니다.")
        self.driver = crawler.setup_driver()
        self.driver.get(board_url)
        self.messages.put(("driver_opened", None))
        self.messages.put(("status", "브라우저 열림 - 로그인 후 수집 시작"))
        self._thread_log("브라우저에서 직접 로그인한 뒤 게시판 목록이 보이면 '수집 시작'을 누르세요.")

    def start_collect(self) -> None:
        if self.driver is None:
            messagebox.showerror("안내", "먼저 브라우저를 열어 주세요.")
            return

        board_url = self.board_url_var.get().strip()
        output_path = self.output_path_var.get().strip()
        start_raw = self.start_page_var.get().strip()
        end_raw = self.end_page_var.get().strip()

        if not output_path:
            messagebox.showerror("입력 오류", "저장 파일 위치를 입력하세요.")
            return
        if Path(output_path).suffix.lower() != ".docx":
            output_path += ".docx"
            self.output_path_var.set(output_path)

        self._run_worker(
            lambda: self._collect_task(board_url, output_path, start_raw, end_raw),
            "수집 중",
        )

    def _collect_task(self, board_url: str, output_path: str, start_raw: str, end_raw: str) -> None:
        crawler.BOARD_URL = board_url
        crawler.OUTPUT_FILE = output_path
        crawler.LOGIN_NOTICE_HANDLER = self._raise_login_notice
        crawler.log = self._thread_log

        try:
            self._ensure_ready_to_collect()
            page_range = self._build_page_range(start_raw, end_raw)
            self._thread_log(f"수집 범위: {page_range.display()} 페이지")

            document_batch = crawler.crawl_board(self.driver, page_range)
            final_files = document_batch.save_final()
            self._thread_log(f"최종 파일: {', '.join(final_files)}")
            self.messages.put(("info", f"수집이 완료되었습니다.\n\n{chr(10).join(final_files)}"))
            self.messages.put(("status", "수집 완료"))
        finally:
            crawler.LOGIN_NOTICE_HANDLER = None

    def _ensure_ready_to_collect(self) -> None:
        self.driver.switch_to.default_content()
        if crawler.page_has_login_or_permission_notice(self.driver):
            raise RuntimeError("현재 화면은 로그인/권한 안내문입니다. 열린 Chrome에서 직접 로그인한 뒤 다시 시도하세요.")

        crawler.switch_to_list_frame(self.driver)

        if crawler.page_has_login_or_permission_notice(self.driver):
            raise RuntimeError("현재 화면은 로그인/권한 안내문입니다. 열린 Chrome에서 직접 로그인한 뒤 다시 시도하세요.")

    def _build_page_range(self, start_raw: str, end_raw: str) -> crawler.PageRange:
        current_page = crawler.get_current_page_number(self.driver) or 1
        last_page = crawler.get_last_page_number(self.driver)

        start_page = self._parse_page_number(start_raw, "시작 페이지") or current_page
        end_page = self._parse_page_number(end_raw, "종료 페이지")
        if end_page is None:
            end_page = last_page

        if last_page is not None and start_page > last_page:
            raise ValueError(f"시작 페이지가 마지막 페이지({last_page})보다 큽니다.")
        if end_page is not None and start_page > end_page:
            raise ValueError("시작 페이지는 종료 페이지보다 클 수 없습니다.")

        return crawler.PageRange(start_page=start_page, end_page=end_page)

    @staticmethod
    def _parse_page_number(raw_value: str, field_name: str) -> int | None:
        value = raw_value.strip()
        if not value:
            return None
        if not value.isdigit():
            raise ValueError(f"{field_name}는 숫자만 입력해야 합니다.")
        page_number = int(value)
        if page_number < 1:
            raise ValueError(f"{field_name}는 1 이상이어야 합니다.")
        return page_number

    @staticmethod
    def _raise_login_notice(_driver) -> None:
        raise RuntimeError("수집 중 로그인 또는 게시판 권한 안내문이 표시되어 중단했습니다.")

    def close_browser(self) -> None:
        if self.driver is None:
            return
        try:
            self.driver.quit()
        except Exception:
            pass
        self.driver = None
        self.messages.put(("driver_closed", None))
        self.messages.put(("status", "브라우저 닫힘"))
        self._append_log("브라우저를 닫았습니다.")

    def on_close(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            if not messagebox.askyesno("종료 확인", "작업이 진행 중입니다. 창을 닫을까요?"):
                return
        self.close_browser()
        self.root.destroy()

    def _run_worker(self, target, status: str) -> None:
        if self.worker is not None and self.worker.is_alive():
            messagebox.showinfo("안내", "현재 작업이 진행 중입니다.")
            return

        self.messages.put(("busy", True))
        self.messages.put(("status", status))

        def wrapped() -> None:
            try:
                target()
            except Exception as exc:
                self._thread_log(traceback.format_exc())
                self.messages.put(("error", str(exc)))
                self.messages.put(("status", "오류 발생"))
            finally:
                self.messages.put(("busy", False))

        self.worker = threading.Thread(target=wrapped, daemon=True)
        self.worker.start()

    def _thread_log(self, message: str) -> None:
        self.messages.put(("log", message))

    def _append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _set_busy(self, busy: bool) -> None:
        if busy:
            self.open_button.configure(state="disabled")
            self.collect_button.configure(state="disabled")
            self.close_button.configure(state="disabled")
            return

        self.open_button.configure(state="disabled" if self.driver is not None else "normal")
        self.collect_button.configure(state="normal" if self.driver is not None else "disabled")
        self.close_button.configure(state="normal" if self.driver is not None else "disabled")

    def _poll_messages(self) -> None:
        while True:
            try:
                kind, payload = self.messages.get_nowait()
            except queue.Empty:
                break

            if kind == "log":
                self._append_log(str(payload))
            elif kind == "status":
                self.status_var.set(str(payload))
            elif kind == "busy":
                self._set_busy(bool(payload))
            elif kind == "driver_opened":
                self._set_busy(False)
            elif kind == "driver_closed":
                self._set_busy(False)
            elif kind == "error":
                messagebox.showerror("오류", str(payload))
            elif kind == "info":
                messagebox.showinfo("완료", str(payload))

        self.root.after(150, self._poll_messages)


def main() -> None:
    if crawler is None:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "실행 오류",
            "필요한 패키지를 불러오지 못해 GUI를 시작할 수 없습니다.\n\n"
            f"{IMPORT_ERROR}\n\n"
            "PowerShell에서 아래 명령을 먼저 실행하세요.\n"
            ".\\.venv\\Scripts\\python.exe -m pip install -r requirements.txt",
        )
        root.destroy()
        return

    root = tk.Tk()
    CafeCollectorGui(root)
    root.mainloop()


if __name__ == "__main__":
    main()
