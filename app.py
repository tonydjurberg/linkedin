import csv
import json
import os
import queue
import re
import threading
import time
from dataclasses import asdict, dataclass
from pathlib import Path
from urllib.parse import parse_qsl, quote_plus, urlencode, urlparse, urlunparse

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import Workbook
from playwright.sync_api import sync_playwright


APP_NAME = "ProspectHunter"
APP_DIR = Path(os.getenv("LOCALAPPDATA", str(Path.home()))) / APP_NAME
PROFILE_DIR = APP_DIR / "ChromeProfile"
STATE_FILE = APP_DIR / "daily_state.json"
DEFAULT_OUTPUT = Path.home() / "Desktop" / "ProspectHunter_Output.xlsx"

LANGUAGES = {
    "English": {"subtitle":"LinkedIn search collector — Chrome only","search":"Search","keywords":"Keywords","title":"Title / role","location":"Location","url":"Exact LinkedIn search URL (optional)","options":"Run options","pages":"Pages","delay":"Delay / page (sec)","visit":"Visit profiles for extra visible details","profiles":"Profiles/day","searches":"Searches/day","output":"Excel output","browse":"Browse","start":"START","pause":"PAUSE","resume":"RESUME","stop":"STOP","running":"RUNNING","finished":"FINISHED","error":"ERROR","today":"Today","seen":"Seen:","new":"New:","language":"Language","readylog":"Ready. Sign in to LinkedIn manually in Chrome when prompted.","starting":"Starting…","finished_msg":"Finished.\\n\\n{} unique contacts exported to:\\n{}","output_log":"Output: {} ({} unique)"},
    "Svenska": {"subtitle":"LinkedIn-sökning — endast Chrome","search":"Sök","keywords":"Sökord","title":"Titel / roll","location":"Plats","url":"Exakt LinkedIn-sök-URL (valfritt)","options":"Köralternativ","pages":"Sidor","delay":"Fördröjning / sida (sek)","visit":"Besök profiler för extra synliga uppgifter","profiles":"Profiler/dag","searches":"Sökningar/dag","output":"Excel-utdata","browse":"Bläddra","start":"START","pause":"PAUS","resume":"FORTSÄTT","stop":"STOPP","running":"KÖR","finished":"KLAR","error":"FEL","today":"Idag","seen":"Sedda:","new":"Nya:","language":"Språk","readylog":"Klar. Logga in manuellt på LinkedIn i Chrome när du blir ombedd.","starting":"Startar…","finished_msg":"Klart.\\n\\n{} unika kontakter exporterades till:\\n{}","output_log":"Utdata: {} ({} unika)"},
    "Português (Brasil)": {"subtitle":"Coletor de buscas do LinkedIn — somente Chrome","search":"Pesquisa","keywords":"Palavras-chave","title":"Cargo / função","location":"Localização","url":"URL exata da busca do LinkedIn (opcional)","options":"Opções de execução","pages":"Páginas","delay":"Atraso / página (seg)","visit":"Visitar perfis para detalhes visíveis extras","profiles":"Perfis/dia","searches":"Pesquisas/dia","output":"Saída Excel","browse":"Procurar","start":"INICIAR","pause":"PAUSAR","resume":"CONTINUAR","stop":"PARAR","running":"EXECUTANDO","finished":"CONCLUÍDO","error":"ERRO","today":"Hoje","seen":"Vistos:","new":"Novos:","language":"Idioma","readylog":"Pronto. Faça login manualmente no LinkedIn pelo Chrome quando solicitado.","starting":"Iniciando…","finished_msg":"Concluído.\\n\\n{} contatos únicos exportados para:\\n{}","output_log":"Saída: {} ({} únicos)"}
}



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


def clean(value):
    return re.sub(r"\s+", " ", value or "").strip()


def unique_lines(value):
    seen, result = set(), []
    for line in (value or "").splitlines():
        line = clean(line)
        if line and line.lower() not in seen:
            seen.add(line.lower())
            result.append(line)
    return result


class DailyLimit:
    def __init__(self, max_profiles=20, max_searches=5):
        self.max_profiles = max_profiles
        self.max_searches = max_searches
        self.day = time.strftime("%Y-%m-%d")
        self.profiles = 0
        self.searches = 0
        self.load()

    def load(self):
        try:
            data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
            if data.get("day") == self.day:
                self.profiles = int(data.get("profiles", 0))
                self.searches = int(data.get("searches", 0))
        except Exception:
            pass

    def save(self):
        APP_DIR.mkdir(parents=True, exist_ok=True)
        STATE_FILE.write_text(json.dumps({
            "day": self.day,
            "profiles": self.profiles,
            "searches": self.searches
        }), encoding="utf-8")

    def can_search(self):
        return self.searches < self.max_searches

    def can_profile(self):
        return self.profiles < self.max_profiles

    def register_search(self):
        self.searches += 1
        self.save()

    def register_profile(self):
        self.profiles += 1
        self.save()


class ProspectHunter:
    def __init__(self, ui_queue):
        self.ui = ui_queue
        self.thread = None
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()

    def log(self, message):
        self.ui.put(("log", message))

    def limits(self, limit):
        self.last_profiles = limit.profiles
        self.last_searches = limit.searches
        self.ui.put(("limits", limit.profiles, limit.max_profiles, limit.searches, limit.max_searches))

    def build_url(self, search_url, keywords, title, location):
        if search_url.strip():
            return search_url.strip()
        query = " ".join(x.strip() for x in (keywords, title, location) if x.strip())
        if not query:
            raise ValueError("Enter keywords/title/location or paste a LinkedIn search URL.")
        return "https://www.linkedin.com/search/results/people/?keywords=" + quote_plus(query)

    def page_url(self, base_url, page_no):
        parts = urlparse(base_url)
        params = dict(parse_qsl(parts.query, keep_blank_values=True))
        params["page"] = str(page_no)
        return urlunparse((parts.scheme, parts.netloc, parts.path, parts.params,
                           urlencode(params), parts.fragment))

    def launch_chrome(self, pw):
        APP_DIR.mkdir(parents=True, exist_ok=True)
        errors = []
        try:
            return pw.chromium.launch_persistent_context(
                user_data_dir=str(PROFILE_DIR),
                channel="chrome",
                headless=False,
                viewport={"width": 1440, "height": 900},
                args=["--start-maximized"],
            )
        except Exception as exc:
            errors.append(str(exc))

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
                except Exception as exc:
                    errors.append(str(exc))
        raise RuntimeError("Could not start Google Chrome. " + " | ".join(errors[-2:]))

    def wait_for_access(self, page):
        self.log("Chrome started. Complete LinkedIn sign-in/security checks in Chrome if requested.")
        deadline = time.time() + 900
        while time.time() < deadline and not self.stop_event.is_set():
            url = page.url.lower()
            if "linkedin.com" in url and not any(x in url for x in ("/login", "/checkpoint", "challenge")):
                return True
            time.sleep(1)
        return False

    def find_result_links(self, page):
        selectors = [
            "main a[href*='/in/']",
            "div[role='main'] a[href*='/in/']",
            "a[href*='/in/']",
        ]
        for selector in selectors:
            try:
                loc = page.locator(selector)
                count = loc.count()
                if count:
                    return [loc.nth(i) for i in range(min(count, 200))]
            except Exception:
                pass
        return []

    def extract_lead_from_link(self, link):
        try:
            href = link.get_attribute("href") or ""
            if "/in/" not in href:
                return None
            href = "https://www.linkedin.com" + href if href.startswith("/") else href
            href = href.split("?")[0].rstrip("/")

            name = clean(link.inner_text(timeout=800))
            if not name:
                try:
                    name = clean(link.get_attribute("aria-label") or "")
                except Exception:
                    pass

            container = None
            for xpath in (
                "xpath=ancestor::li[1]",
                "xpath=ancestor::article[1]",
                "xpath=ancestor::*[@role='listitem'][1]",
            ):
                try:
                    candidate = link.locator(xpath)
                    if candidate.count():
                        container = candidate
                        break
                except Exception:
                    pass

            text = clean(container.inner_text(timeout=1200)) if container else name
            lines = unique_lines(container.inner_text(timeout=1200) if container else name)

            headline = ""
            location = ""
            snippet = ""
            if container:
                for selector in (
                    ".entity-result__primary-subtitle",
                    "[data-anonymize='job-title']",
                    "div.t-14.t-black.t-normal",
                ):
                    try:
                        value = clean(container.locator(selector).first.inner_text(timeout=500))
                        if value:
                            headline = value
                            break
                    except Exception:
                        pass
                for selector in (
                    ".entity-result__secondary-subtitle",
                    "[data-anonymize='location']",
                    "div.t-14.t-normal",
                ):
                    try:
                        value = clean(container.locator(selector).first.inner_text(timeout=500))
                        if value:
                            location = value
                            break
                    except Exception:
                        pass
                for selector in (
                    ".entity-result__summary",
                    ".search-result__snippets",
                    "[data-anonymize='summary']",
                ):
                    try:
                        value = clean(container.locator(selector).first.inner_text(timeout=500))
                        if value:
                            snippet = value
                            break
                    except Exception:
                        pass

            if not name and lines:
                name = lines[0]
            if not headline and len(lines) > 1:
                headline = lines[1]
            if not location and len(lines) > 2:
                location = lines[-1]

            company = ""
            try:
                company = clean(container.locator("[data-anonymize='company-name']").first.inner_text(timeout=400))
            except Exception:
                pass

            return Lead(
                name=name,
                headline=headline,
                location=location,
                company=company,
                profile_url=href,
                snippet=snippet or text,
                collected_at=time.strftime("%Y-%m-%d %H:%M:%S"),
            )
        except Exception:
            return None

    def collect_page(self, page, records, limit):
        links = self.find_result_links(page)
        self.log(f"Detected {len(links)} LinkedIn profile links on this page.")
        added = 0
        for link in links:
            if self.stop_event.is_set() or not limit.can_profile():
                break
            while self.pause_event.is_set() and not self.stop_event.is_set():
                time.sleep(0.5)
            lead = self.extract_lead_from_link(link)
            if not lead or not lead.profile_url:
                continue
            key = lead.profile_url.lower()
            if key in records:
                continue
            records[key] = lead
            limit.register_profile()
            added += 1
            self.ui.put(("count", len(records), added))
            self.limits(limit)
        return len(links), added

    def enrich_profile(self, page, lead):
        try:
            page.goto(lead.profile_url, wait_until="domcontentloaded", timeout=30000)
            time.sleep(1.5)
            if any(x in page.url.lower() for x in ("/login", "/checkpoint", "challenge")):
                return
            if not lead.name:
                title = clean(page.title())
                lead.name = title.split("|")[0].strip()
            main = page.locator("main")
            body = clean(main.inner_text(timeout=2000))
            candidates = []
            for selector in ("div.text-body-medium", "main h2", "[data-anonymize='headline']"):
                try:
                    for i in range(min(main.locator(selector).count(), 5)):
                        value = clean(main.locator(selector).nth(i).inner_text(timeout=500))
                        if value:
                            candidates.append(value)
                except Exception:
                    pass
            if not lead.headline:
                for value in candidates:
                    if value.lower() != lead.name.lower() and len(value) < 180:
                        lead.headline = value
                        break
            if not lead.location:
                for line in unique_lines(body):
                    low = line.lower()
                    if any(word in low for word in (
                        "sweden", "sverige", "stockholm", "gothenburg", "göteborg",
                        "malmö", "skåne", "brazil", "brasil"
                    )):
                        lead.location = line
                        break
        except Exception:
            pass

    def run(self, settings):
        self.stop_event.clear()
        self.pause_event.clear()
        records = {}
        limit = DailyLimit(settings["max_profiles"], settings["max_searches"])
        self.limits(limit)
        context = None

        try:
            target = self.build_url(
                settings["search_url"], settings["keywords"],
                settings["title"], settings["location"]
            )
            pages = max(1, int(settings["pages"]))

            with sync_playwright() as pw:
                context = self.launch_chrome(pw)
                page = context.pages[0] if context.pages else context.new_page()

                self.log("Opening LinkedIn search…")
                page.goto(target, wait_until="domcontentloaded", timeout=60000)
                if not self.wait_for_access(page):
                    self.log("Stopped waiting for LinkedIn access.")
                    return

                for page_no in range(1, pages + 1):
                    if self.stop_event.is_set():
                        break
                    while self.pause_event.is_set() and not self.stop_event.is_set():
                        time.sleep(0.5)

                    if not limit.can_search():
                        self.log(f"Daily search limit reached: {limit.searches}/{limit.max_searches}.")
                        break

                    if page_no > 1:
                        page.goto(self.page_url(target, page_no),
                                  wait_until="domcontentloaded", timeout=60000)

                    limit.register_search()
                    self.limits(limit)
                    time.sleep(max(1.0, float(settings["delay"])))

                    if any(x in page.url.lower() for x in ("/checkpoint", "challenge")):
                        self.log("LinkedIn security verification detected. Complete it in Chrome; scraping is paused.")
                        while any(x in page.url.lower() for x in ("/checkpoint", "challenge")) and not self.stop_event.is_set():
                            time.sleep(1)

                    before = len(records)
                    link_count, added = self.collect_page(page, records, limit)
                    self.log(f"Page {page_no}: {added} new contacts, {len(records)} unique total.")

                    if link_count == 0:
                        self.log("No LinkedIn profile links found; stopping pagination.")
                        break
                    if not limit.can_profile():
                        self.log(f"Daily profile limit reached: {limit.profiles}/{limit.max_profiles}.")
                        break

                if settings["visit_profiles"] and records:
                    self.log("Enriching visible profile details in the same Chrome tab…")
                    for index, lead in enumerate(list(records.values()), 1):
                        if self.stop_event.is_set():
                            break
                        while self.pause_event.is_set() and not self.stop_event.is_set():
                            time.sleep(0.5)
                        self.log(f"Profile {index}/{len(records)}: {lead.name or lead.profile_url}")
                        self.enrich_profile(page, lead)

                leads = list(records.values())
                self.export(leads, settings["output"])
                self.log(f"FINISHED. {len(leads)} unique contacts exported.")
                self.ui.put(("done", str(settings["output"]), len(leads)))

        except Exception as exc:
            self.log("ERROR: " + str(exc))
            self.ui.put(("error", str(exc)))
        finally:
            if context:
                try:
                    context.close()
                except Exception:
                    pass

    def export(self, leads, output_path):
        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        fields = list(asdict(Lead()).keys())

        csv_path = output.with_suffix(".csv")
        with csv_path.open("w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.DictWriter(handle, fieldnames=fields)
            writer.writeheader()
            for lead in leads:
                writer.writerow(asdict(lead))

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Leads"
        sheet.append(fields)
        for lead in leads:
            row = asdict(lead)
            sheet.append([row[field] for field in fields])
        sheet.freeze_panes = "A2"
        for column in sheet.columns:
            width = min(60, max(12, max(len(str(cell.value or "")) for cell in column) + 2))
            sheet.column_dimensions[column[0].column_letter].width = width
        workbook.save(output)

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
        self.max_profiles = tk.IntVar(value=20)
        self.max_searches = tk.IntVar(value=5)
        self.status = tk.StringVar(value="Ready")
        self.language = tk.StringVar(value="English")
        self.limit_status = tk.StringVar(value="Today: 0/20 profiles | 0/5 searches")
        self.found = tk.IntVar(value=0)
        self.new = tk.IntVar(value=0)
        self.build_ui()
        self.root.after(250, self.poll)

    def t(self, key):
        return LANGUAGES[self.language.get()].get(key, key)

    def build_ui(self):
        for child in self.root.winfo_children():
            child.destroy()

        outer = ttk.Frame(self.root, padding=16)
        outer.pack(fill="both", expand=True)
        header = ttk.Frame(outer)
        header.pack(fill="x")
        ttk.Label(header, text="PROSPECT HUNTER", font=("Segoe UI", 18, "bold")).pack(side="left")
        lang_frame = ttk.Frame(header)
        lang_frame.pack(side="right")
        ttk.Label(lang_frame, text=self.t("language")).pack(side="left", padx=(0, 6))
        ttk.Combobox(lang_frame, textvariable=self.language,
                     values=list(LANGUAGES.keys()), state="readonly", width=20).pack(side="left")
        self.language.trace_add("write", self.change_language)

        self.subtitle_label = ttk.Label(outer, text=self.t("subtitle"))
        self.subtitle_label.pack(anchor="w", pady=(0, 12))

        form = ttk.LabelFrame(outer, text=self.t("search"))
        form.pack(fill="x")
        self.form = form
        self.form_labels = []
        rows = [("keywords", self.keywords), ("title", self.title_var),
                ("location", self.location), ("url", self.search_url)]
        for row, (key, var) in enumerate(rows):
            label = ttk.Label(form, text=self.t(key), width=30)
            label.grid(row=row, column=0, sticky="w", padx=8, pady=6)
            self.form_labels.append((key, label))
            ttk.Entry(form, textvariable=var).grid(row=row, column=1, sticky="ew", padx=8, pady=6)
        form.columnconfigure(1, weight=1)

        opts = ttk.LabelFrame(outer, text=self.t("options"))
        opts.pack(fill="x", pady=12)
        self.opts = opts
        self.option_labels = []
        def ol(key, row, col):
            w = ttk.Label(opts, text=self.t(key))
            w.grid(row=row, column=col, padx=8, pady=6, sticky="w")
            self.option_labels.append((key, w))
        ol("pages", 0, 0)
        ttk.Spinbox(opts, from_=1, to=500, textvariable=self.pages, width=8).grid(row=0, column=1, padx=8, pady=6, sticky="w")
        ol("delay", 0, 2)
        ttk.Spinbox(opts, from_=1, to=60, increment=1, textvariable=self.delay, width=8).grid(row=0, column=3, padx=8, pady=6, sticky="w")
        self.visit_check = ttk.Checkbutton(opts, text=self.t("visit"), variable=self.visit_profiles)
        self.visit_check.grid(row=0, column=4, padx=12, pady=6, sticky="w")
        ol("profiles", 1, 0)
        ttk.Spinbox(opts, from_=1, to=5000, textvariable=self.max_profiles, width=8).grid(row=1, column=1, padx=8, pady=6, sticky="w")
        ol("searches", 1, 2)
        ttk.Spinbox(opts, from_=1, to=500, textvariable=self.max_searches, width=8).grid(row=1, column=3, padx=8, pady=6, sticky="w")

        out = ttk.Frame(outer)
        out.pack(fill="x", pady=(0, 10))
        self.output_label = ttk.Label(out, text=self.t("output"))
        self.output_label.pack(side="left")
        ttk.Entry(out, textvariable=self.output).pack(side="left", fill="x", expand=True, padx=8)
        self.browse_button = ttk.Button(out, text=self.t("browse"), command=self.browse)
        self.browse_button.pack(side="left")

        buttons = ttk.Frame(outer)
        buttons.pack(fill="x", pady=6)
        self.start_button = ttk.Button(buttons, text=self.t("start"), command=self.start)
        self.start_button.pack(side="left", padx=(0, 6))
        self.pause_button = ttk.Button(buttons, text=self.t("pause"), command=self.engine.pause)
        self.pause_button.pack(side="left", padx=6)
        self.resume_button = ttk.Button(buttons, text=self.t("resume"), command=self.engine.resume)
        self.resume_button.pack(side="left", padx=6)
        self.stop_button = ttk.Button(buttons, text=self.t("stop"), command=self.engine.stop)
        self.stop_button.pack(side="left", padx=6)

        stats = ttk.Frame(outer)
        stats.pack(fill="x", pady=6)
        ttk.Label(stats, textvariable=self.status, font=("Segoe UI", 10, "bold")).pack(side="left")
        self.limit_status_label = ttk.Label(stats, textvariable=self.limit_status)
        self.limit_status_label.pack(side="left", padx=(20, 0))
        self.seen_label = ttk.Label(stats, text=self.t("seen"))
        self.seen_label.pack(side="left")
        ttk.Label(stats, textvariable=self.found).pack(side="left")
        self.new_label = ttk.Label(stats, text=self.t("new"))
        self.new_label.pack(side="left")
        ttk.Label(stats, textvariable=self.new).pack(side="left")

        self.log = tk.Text(outer, height=20, wrap="word")
        self.log.pack(fill="both", expand=True, pady=(8, 0))
        self.log.insert("end", self.t("readylog") + "\n")
        self.log.configure(state="disabled")

    def change_language(self, *_):
        if not hasattr(self, "subtitle_label"):
            return
        self.root.title(APP_NAME + " — " + self.language.get())
        self.subtitle_label.config(text=self.t("subtitle"))
        self.form.config(text=self.t("search"))
        self.opts.config(text=self.t("options"))
        for key, widget in self.form_labels + self.option_labels:
            widget.config(text=self.t(key))
        self.visit_check.config(text=self.t("visit"))
        self.output_label.config(text=self.t("output"))
        self.browse_button.config(text=self.t("browse"))
        self.start_button.config(text=self.t("start"))
        self.pause_button.config(text=self.t("pause"))
        self.resume_button.config(text=self.t("resume"))
        self.stop_button.config(text=self.t("stop"))
        self.seen_label.config(text=self.t("seen"))
        self.new_label.config(text=self.t("new"))
        self.update_limit_text()

    def update_limit_text(self):
        profiles = getattr(self.engine, "last_profiles", 0)
        searches = getattr(self.engine, "last_searches", 0)
        self.limit_status.set(
            f"{self.t('today')}: {profiles}/{self.max_profiles.get()} profiles | "
            f"{searches}/{self.max_searches.get()} searches"
        )

    def browse(self):
        path = filedialog.asksaveasfilename(
            title="Save Excel output",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="ProspectHunter_Output.xlsx",
        )
        if path:
            self.output.set(path)

    def add_log(self, message):
        self.log.configure(state="normal")
        self.log.insert("end", message + "\n")
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
                "max_profiles": int(self.max_profiles.get()),
                "max_searches": int(self.max_searches.get()),
            }
            self.status.set(self.t("running"))
            self.found.set(0)
            self.new.set(0)
            self.add_log(self.t("starting"))
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
                elif item[0] == "limits":
                    self.engine.last_profiles = item[1]
                    self.engine.last_searches = item[3]
                    self.update_limit_text()
                elif item[0] == "done":
                    self.status.set(self.t("finished"))
                    self.add_log(self.t("output_log").format(item[1], item[2]))
                    messagebox.showinfo(APP_NAME, self.t("finished_msg").format(item[2], item[1]))
                elif item[0] == "error":
                    self.status.set(self.t("error"))
                    messagebox.showerror(APP_NAME, item[1])
                self.q.task_done()
        except queue.Empty:
            pass
        self.root.after(250, self.poll)


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
