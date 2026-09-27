import csv
import os
import queue
import re
import threading
import time
import webbrowser
from dataclasses import dataclass, asdict
from pathlib import Path
from urllib.parse import quote_plus, urlparse, parse_qsl, urlencode, urlunparse

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from openpyxl import Workbook
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


APP_NAME = "ProspectHunter"
APP_DIR = Path(os.getenv("LOCALAPPDATA", Path.home())) / APP_NAME
PROFILE_DIR = APP_DIR / "ChromeProfile"
DEFAULT_OUTPUT = Path.home() / "Desktop" / "ProspectHunter_Output.xlsx"


@dataclass
class Lead:
    name: str = ""
    headline: str = ""
    location: str = ""
    company: str = ""
    profile_url: str = ""
    snippet: str = ""
    source: str = "LinkedIn"
    collected_at: str = ""


def clean(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def unique_lines(text: str):
    out = []
    seen = set()
    for raw in (text or "").splitlines():
        value = clean(raw)
        if value and value.lower() not in seen:
            seen.add(value.lower())
            out.append(value)
    return out


class ProspectHunter:
    def __init__(self, ui_queue: queue.Queue):
        self.ui = ui_queue
        self.thread = None
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()

    def log(self, message):
        self.ui.put(("log", message))

    def count(self, found, new):
        self.ui.put(("count", found, new))

    def build_url(self, search_url, keywords, title, location):
        if search_url.strip():
            return search_url.strip()
        parts = [keywords.strip(), title.strip(), location.strip()]
        query = " ".join(p for p in parts if p)
        if not query:
            raise ValueError("Enter keywords/title/location or paste a LinkedIn search URL.")
        return "https://www.linkedin.com/search/results/people/?keywords=" + quote_plus(query)

    def next_page_url(self, url, page_no):
        parts = urlparse(url)
        params = dict(parse_qsl(parts.query, keep_blank_values=True))
        params["page"] = str(page_no)
        return urlunparse((parts.scheme, parts.netloc, parts.path, parts.params, urlencode(params), parts.fragment))

    def launch_chrome(self, pw):
        APP_DIR.mkdir(parents=True, exist_ok=True)
        try:
            return pw.chromium.launch_persistent_context(
                user_data_dir=str(PROFILE_DIR),
                channel="chrome",
                headless=False,
                viewport={"width": 1440, "height": 900},
                args=["--start-maximized"],
            )
        except Exception as first_error:
            candidates = [
                os.path.expandvars(r"%PROGRAMFILES%\Google\Chrome\Application\chrome.exe"),
                os.path.expandvars(r"%PROGRAMFILES(X86)%\Google\Chrome\Application\chrome.exe"),
                os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
            ]
            for exe in candidates:
                if os.path.exists(exe):
                    try:
                        return pw.chromium.launch_persistent_context(
                            user_data_dir=str(PROFILE_DIR),
                            executable_path=exe,
                            headless=False,
                            viewport={"width": 1440, "height": 900},
                            args=["--start-maximized"],
                        )
                    except Exception:
                        pass
            raise RuntimeError(f"Could not start Google Chrome. {first_error}")

    def wait_for_access(self, page):
        self.log("Chrome started. If LinkedIn asks for sign-in or verification, complete it in Chrome; the program will wait.")
        deadline = time.time() + 900
        while time.time() < deadline and not self.stop_event.is_set():
            url = page.url.lower()
            if "linkedin.com" in url and "/login" not in url and "/checkpoint" not in url and "challenge" not in url:
                return True
            time.sleep(1)
        return False

    def extract_cards(self, page):
        selectors = [
            "li.reusable-search__result-container",
            "div[data-view-name='search-entity-result-universal-template']",
            "div.search-result__wrapper",
        ]
        cards = []
        for selector in selectors:
            try:
                loc = page.locator(selector)
                if loc.count() > 0:
                    cards = [loc.nth(i) for i in range(min(loc.count(), 200))]
                    break
            except Exception:
                continue
        if not cards:
            # Fallback: profile links grouped by nearest li/article/section.
            links = page.locator("a[href*='/in/']")
            n = min(links.count(), 200)
            for i in range(n):
                try:
                    cards.append(links.nth(i).locator("xpath=ancestor::*[self::li or self::article or @role='listitem'][1]"))
                except Exception:
                    pass
        return cards

    def extract_lead(self, card):
        try:
            links = card.locator("a[href*='/in/']")
            if links.count() == 0:
                return None
            href = links.first.get_attribute("href") or ""
            if not href:
                return None
            href = "https://www.linkedin.com" + href if href.startswith("/") else href.split("?")[0]

            name = ""
            for sel in [
                ".entity-result__title-text a",
                ".entity-result__title-text",
                "a[href*='/in/'] span",
                "a[href*='/in/']",
            ]:
                try:
                    value = clean(card.locator(sel).first.inner_text(timeout=800))
                    if value:
                        name = value
                        break
                except Exception:
                    pass

            headline = ""
            location = ""
            company = ""
            snippet = ""
            for sel in [".entity-result__primary-subtitle", "div.t-14.t-black.t-normal"]:
                try:
                    headline = clean(card.locator(sel).first.inner_text(timeout=500))
                    if headline:
                        break
                except Exception:
                    pass
            for sel in [".entity-result__secondary-subtitle", "div.t-14.t-normal"]:
                try:
                    location = clean(card.locator(sel).first.inner_text(timeout=500))
                    if location:
                        break
                except Exception:
                    pass
            for sel in [".entity-result__summary", ".search-result__snippets"]:
                try:
                    snippet = clean(card.locator(sel).first.inner_text(timeout=500))
                    if snippet:
                        break
                except Exception:
                    pass

            lines = unique_lines(card.inner_text(timeout=1000))
            if lines:
                if not name:
                    name = lines[0]
                if not headline and len(lines) > 1:
                    headline = lines[1]
                if not location and len(lines) > 2:
                    location = lines[-1]

            # Best-effort company extraction from common entity metadata.
            for sel in [".entity-result__primary-subtitle", "[data-anonymize='company-name']"]:
                try:
                    txt = clean(card.locator(sel).first.inner_text(timeout=400))
                    if txt and company == "":
                        company = txt
                except Exception:
                    pass

            return Lead(name=name, headline=headline, location=location, company=company,
                        profile_url=href, snippet=snippet, collected_at=time.strftime("%Y-%m-%d %H:%M:%S"))
        except Exception:
            return None

    def enrich_profile(self, page, lead: Lead):
        try:
            page.goto(lead.profile_url, wait_until="domcontentloaded", timeout=30000)
            time.sleep(1.5)
            if "/login" in page.url.lower() or "/checkpoint" in page.url.lower():
                return
            title = clean(page.title())
            if title and not lead.name:
                lead.name = title.split("|")[0].strip()

            candidates = []
            for sel in [
                "div.text-body-medium",
                "div.pv-text-details__left-panel div.text-body-medium",
                "main h2",
                "main section div.text-body-medium",
            ]:
                try:
                    n = page.locator(sel).count()
                    for i in range(min(n, 5)):
                        txt = clean(page.locator(sel).nth(i).inner_text(timeout=500))
                        if txt:
                            candidates.append(txt)
                except Exception:
                    pass
            for c in candidates:
                if c != lead.name and len(c) < 160:
                    if not lead.headline:
                        lead.headline = c
                    break

            try:
                body = clean(page.locator("main").inner_text(timeout=1500))
                if not lead.location:
                    lines = unique_lines(body)
                    for line in lines:
                        if any(x in line.lower() for x in ["sweden", "stockholm", "gothenburg", "malmö", "skåne", "brazil", "sverige"]):
                            lead.location = line
                            break
            except Exception:
                pass
        except Exception:
            pass

    def run(self, settings):
        self.stop_event.clear()
        self.pause_event.clear()
        found = new = 0
        records = {}
        try:
            target = self.build_url(settings["search_url"], settings["keywords"], settings["title"], settings["location"])
            max_pages = max(1, int(settings["pages"]))
            visit_profiles = settings["visit_profiles"]

            with sync_playwright() as pw:
                context = self.launch_chrome(pw)
                page = context.pages[0] if context.pages else context.new_page()
                self.log("Opening LinkedIn search…")
                page.goto(target, wait_until="domcontentloaded", timeout=60000)

                if not self.wait_for_access(page):
                    self.log("Stopped waiting for LinkedIn access.")
                    context.close()
                    return

                for page_no in range(1, max_pages + 1):
                    if self.stop_event.is_set():
                        break
                    while self.pause_event.is_set() and not self.stop_event.is_set():
                        time.sleep(0.5)

                    if page_no > 1:
                        page.goto(self.next_page_url(target, page_no), wait_until="domcontentloaded", timeout=60000)

                    time.sleep(max(1.0, settings["delay"]))
                    if "checkpoint" in page.url.lower() or "challenge" in page.url.lower():
                        self.log("LinkedIn security verification detected. Complete it in Chrome; scraping is paused.")
                        while ("checkpoint" in page.url.lower() or "challenge" in page.url.lower()) and not self.stop_event.is_set():
                            time.sleep(1)
                        if self.stop_event.is_set():
                            break

                    cards = self.extract_cards(page)
                    self.log(f"Page {page_no}: {len(cards)} result cards detected.")
                    before = len(records)

                    for card in cards:
                        if self.stop_event.is_set():
                            break
                        while self.pause_event.is_set() and not self.stop_event.is_set():
                            time.sleep(0.5)
                        lead = self.extract_lead(card)
                        if not lead or not lead.profile_url:
                            continue
                        key = lead.profile_url.rstrip("/").lower()
                        if key not in records:
                            records[key] = lead
                            new += 1
                        found += 1
                        self.count(found, new)

                    self.log(f"Page {page_no}: {len(records)-before} new contacts.")
                    if not cards:
                        self.log("No result cards found; stopping pagination.")
                        break

                if visit_profiles and records:
                    profile_page = context.new_page()
                    for idx, lead in enumerate(list(records.values()), 1):
                        if self.stop_event.is_set():
                            break
                        while self.pause_event.is_set() and not self.stop_event.is_set():
                            time.sleep(0.5)
                        self.log(f"Enriching profile {idx}/{len(records)}: {lead.name or lead.profile_url}")
                        self.enrich_profile(profile_page, lead)
                    profile_page.close()

                context.close()

            leads = list(records.values())
            self.export(leads, settings["output"])
            self.log(f"FINISHED. {len(leads)} unique contacts exported.")
            self.ui.put(("done", str(settings["output"]), len(leads)))
        except Exception as exc:
            self.log(f"ERROR: {exc}")
            self.ui.put(("error", str(exc)))

    def export(self, leads, output_path):
        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        csv_path = output.with_suffix(".csv")
        with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=list(asdict(Lead()).keys()))
            writer.writeheader()
            for lead in leads:
                writer.writerow(asdict(lead))

        wb = Workbook()
        ws = wb.active
        ws.title = "Leads"
        headers = list(asdict(Lead()).keys())
        ws.append(headers)
        for lead in leads:
            row = asdict(lead)
            ws.append([row[h] for h in headers])
        ws.freeze_panes = "A2"
        for col in ws.columns:
            width = min(60, max(12, max(len(str(c.value or "")) for c in col) + 2))
            ws.column_dimensions[col[0].column_letter].width = width
        wb.save(output)

    def start(self, settings):
        if self.thread and self.thread.is_alive():
            return
        self.thread = threading.Thread(target=self.run, args=(settings,), daemon=True)
        self.thread.start()

    def pause(self):
        self.pause_event.set()

    def resume(self):
        self.pause_event.clear()

    def stop(self):
        self.stop_event.set()
        self.pause_event.clear()


class App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("920x720")
        self.root.minsize(820, 620)
        self.q = queue.Queue()
        self.engine = ProspectHunter(self.q)

        self.search_url = tk.StringVar()
        self.keywords = tk.StringVar()
        self.title_var = tk.StringVar()
        self.location = tk.StringVar()
        self.pages = tk.IntVar(value=10)
        self.delay = tk.DoubleVar(value=10.0)
        self.output = tk.StringVar(value=str(DEFAULT_OUTPUT))
        self.visit_profiles = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value="Ready")
        self.found = tk.IntVar(value=0)
        self.new = tk.IntVar(value=0)

        self.build_ui()
        self.root.after(250, self.poll)

    def build_ui(self):
        outer = ttk.Frame(self.root, padding=16)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="PROSPECT HUNTER", font=("Segoe UI", 18, "bold")).pack(anchor="w")
        ttk.Label(outer, text="LinkedIn search collector — Chrome only").pack(anchor="w", pady=(0, 12))

        form = ttk.LabelFrame(outer, text="Search")
        form.pack(fill="x")

        rows = [
            ("Keywords", self.keywords),
            ("Title / role", self.title_var),
            ("Location", self.location),
            ("Exact LinkedIn search URL (optional)", self.search_url),
        ]
        for r, (label, var) in enumerate(rows):
            ttk.Label(form, text=label, width=30).grid(row=r, column=0, sticky="w", padx=8, pady=6)
            ttk.Entry(form, textvariable=var).grid(row=r, column=1, sticky="ew", padx=8, pady=6)
        form.columnconfigure(1, weight=1)

        opts = ttk.LabelFrame(outer, text="Run options")
        opts.pack(fill="x", pady=12)
        ttk.Label(opts, text="Pages").grid(row=0, column=0, padx=8, pady=6, sticky="w")
        ttk.Spinbox(opts, from_=1, to=500, textvariable=self.pages, width=8).grid(row=0, column=1, padx=8, pady=6, sticky="w")
        ttk.Label(opts, text="Delay / page (sec) — default 10").grid(row=0, column=2, padx=8, pady=6, sticky="w")
        ttk.Spinbox(opts, from_=0.5, to=30, increment=0.5, textvariable=self.delay, width=8).grid(row=0, column=3, padx=8, pady=6, sticky="w")
        ttk.Checkbutton(opts, text="Visit profiles for extra visible details", variable=self.visit_profiles).grid(row=0, column=4, padx=12, pady=6, sticky="w")

        out = ttk.Frame(outer)
        out.pack(fill="x", pady=(0, 10))
        ttk.Label(out, text="Excel output").pack(side="left")
        ttk.Entry(out, textvariable=self.output).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(out, text="Browse", command=self.browse).pack(side="left")

        buttons = ttk.Frame(outer)
        buttons.pack(fill="x", pady=6)
        ttk.Button(buttons, text="START", command=self.start).pack(side="left", padx=(0, 6))
        ttk.Button(buttons, text="PAUSE", command=self.engine.pause).pack(side="left", padx=6)
        ttk.Button(buttons, text="RESUME", command=self.engine.resume).pack(side="left", padx=6)
        ttk.Button(buttons, text="STOP", command=self.engine.stop).pack(side="left", padx=6)

        stats = ttk.Frame(outer)
        stats.pack(fill="x", pady=6)
        ttk.Label(stats, textvariable=self.status, font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Label(stats, text="   Seen:").pack(side="left")
        ttk.Label(stats, textvariable=self.found).pack(side="left")
        ttk.Label(stats, text="   New:").pack(side="left")
        ttk.Label(stats, textvariable=self.new).pack(side="left")

        self.log = tk.Text(outer, height=20, wrap="word")
        self.log.pack(fill="both", expand=True, pady=(8, 0))
        self.log.insert("end", "Ready. Sign in to LinkedIn manually in the Chrome window when prompted.\n")
        self.log.configure(state="disabled")

    def browse(self):
        path = filedialog.asksaveasfilename(
            title="Save Excel output",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="ProspectHunter_Output.xlsx",
        )
        if path:
            self.output.set(path)

    def add_log(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def start(self):
        try:
            settings = {
                "search_url": self.search_url.get(),
                "keywords": self.keywords.get(),
                "title": self.title_var.get(),
                "location": self.location.get(),
                "pages": int(self.pages.get()),
                "delay": float(self.delay.get()),
                "output": self.output.get(),
                "visit_profiles": bool(self.visit_profiles.get()),
            }
            self.status.set("RUNNING")
            self.found.set(0)
            self.new.set(0)
            self.add_log("Starting…")
            self.engine.start(settings)
        except Exception as exc:
            messagebox.showerror(APP_NAME, str(exc))

    def poll(self):
        try:
            while True:
                item = self.q.get_nowait()
                if item[0] == "log":
                    self.add_log(item[1])
                elif item[0] == "count":
                    self.found.set(item[1])
                    self.new.set(item[2])
                elif item[0] == "done":
                    self.status.set("FINISHED")
                    self.add_log(f"Output: {item[1]} ({item[2]} unique)")
                    messagebox.showinfo(APP_NAME, f"Finished.\n\n{item[2]} unique contacts exported to:\n{item[1]}")
                elif item[0] == "error":
                    self.status.set("ERROR")
                    messagebox.showerror(APP_NAME, item[1])
                self.q.task_done()
        except queue.Empty:
            pass
        self.root.after(250, self.poll)


if __name__ == "__main__":
    root = tk.Tk()
    try:
        root.iconname(APP_NAME)
    except Exception:
        pass
    App(root)
    root.mainloop()
