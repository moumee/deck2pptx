"""Drives the deck in a real browser: navigation, measurement and isolated captures."""
from __future__ import annotations

import functools
import http.server
import io
import os
import threading
import time
import urllib.parse
from dataclasses import dataclass, field
from pathlib import Path

from PIL import Image
from playwright.sync_api import Browser, Error as PlaywrightError, Page, sync_playwright

from .fonts import CssFace, parse_font_faces

EXTRACT_JS = (Path(__file__).parent / "extract.js").read_text(encoding="utf8")
NAV_TIMEOUT_MS = 45_000


class _QuietHandler(http.server.SimpleHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        super().end_headers()


class _Server(http.server.ThreadingHTTPServer):
    daemon_threads = True

    def handle_error(self, request, client_address):
        pass  # browsers drop connections mid-transfer; that is not a failure


def launch(pw) -> Browser:
    """Prefer a browser the user already has, so nothing extra must be downloaded."""
    errors = []
    for channel in ("chrome", "msedge", None):
        try:
            return pw.chromium.launch(channel=channel, headless=True,
                                      args=["--force-color-profile=srgb", "--font-render-hinting=none"])
        except PlaywrightError as e:
            errors.append(f"{channel or 'bundled chromium'}: {str(e).splitlines()[0]}")
    raise RuntimeError("No Chromium-based browser found. Install Chrome, or run: python -m playwright install chromium\n"
                       + "\n".join(errors))


@dataclass
class Deck:
    count: int
    title: str
    lang: str
    width: float
    height: float
    css_faces: list[CssFace] = field(default_factory=list)
    stylesheet_urls: list[str] = field(default_factory=list)
    css_vars: dict = field(default_factory=dict)


class DeckBrowser:
    def __init__(self, html_path: str | os.PathLike, dpr: int = 2, viewport=(1920, 1080), log=print):
        self.path = Path(html_path).resolve()
        self.dpr = dpr
        self.viewport = viewport
        self.log = log
        self.responses = []
        self._cdp = None

    def __enter__(self) -> "DeckBrowser":
        handler = functools.partial(_QuietHandler, directory=str(self.path.parent))
        self.server = _Server(("127.0.0.1", 0), handler)
        threading.Thread(target=self.server.serve_forever, daemon=True).start()
        self.base = f"http://127.0.0.1:{self.server.server_address[1]}/"
        self.pw = sync_playwright().start()
        self.browser = launch(self.pw)
        self.context = self.browser.new_context(
            viewport={"width": self.viewport[0], "height": self.viewport[1]},
            device_scale_factor=self.dpr, reduced_motion="no-preference")
        self.page: Page = self.context.new_page()
        self.page.on("response", lambda r: self.responses.append(r))
        return self

    def __exit__(self, *exc):
        try:
            self.browser.close()
            self.pw.stop()
        finally:
            self.server.shutdown()
            self.server.server_close()

    def open(self, hints: dict | None = None) -> Deck:
        url = self.base + urllib.parse.quote(self.path.name)
        try:
            self.page.goto(url, wait_until="networkidle", timeout=NAV_TIMEOUT_MS)
        except PlaywrightError as e:
            self.log(f"  page never went idle ({str(e).splitlines()[0]}); continuing with what loaded")
        self.page.evaluate("() => document.fonts.ready")
        self.page.add_script_tag(content=EXTRACT_JS)
        info = self.page.evaluate("() => __d2p.init()")
        if hints:
            self.page.evaluate("""h => {
                for (const s of h.raster || []) document.querySelectorAll(s).forEach(e => e.setAttribute('data-pptx', 'raster'));
                for (const s of h.native || []) document.querySelectorAll(s).forEach(e => e.setAttribute('data-pptx', 'native'));
                for (const s of h.ignore || []) document.querySelectorAll(s).forEach(e => e.setAttribute('data-pptx-ignore', ''));
            }""", hints)
        if not info["count"]:
            raise RuntimeError(f"No slides found in {self.path.name}: expected .slide elements or <deck-stage> children.")
        self.page.evaluate("() => { __d2p.goto(0); __d2p.settle(0); }")
        size = self.page.evaluate("() => { const r = __d2p.slides[0].getBoundingClientRect(); return [r.width, r.height]; }")
        return self._deck(info, size)

    def _deck(self, info: dict, size: list[float]) -> Deck:
        urls, faces = [], []
        for r in self.responses:
            try:
                if r.request.resource_type == "stylesheet" and r.ok:
                    urls.append(r.url)
                    faces += parse_font_faces(r.text(), r.url)
            except PlaywrightError:
                continue
        inline = self.page.evaluate("() => [...document.querySelectorAll('style')].map(s => s.textContent).join('\\n')")
        faces += parse_font_faces(inline, self.page.url)
        return Deck(info["count"], info["title"], info["lang"], size[0], size[1], faces, urls,
                    self.page.evaluate("() => __d2p.cssVars()"))

    def fetch_local(self, url: str) -> bytes | None:
        """Bytes the browser already downloaded, so nothing is fetched twice."""
        for r in self.responses:
            if r.url == url:
                try:
                    return r.body()
                except PlaywrightError:
                    return None
        if url.startswith(self.base):
            p = self.path.parent / urllib.parse.unquote(url[len(self.base):].split("?")[0])
            return p.read_bytes() if p.is_file() else None
        return None

    def used_fonts(self, family: str, weight: str, italic: bool, text: str) -> list[dict]:
        """What Chrome really drew this text with, which a font stack alone cannot tell."""
        self.page.evaluate("a => __d2p.probeFont(a)", {"family": family, "weight": str(weight), "italic": italic, "text": text})
        if self._cdp is None:
            self._cdp = self.context.new_cdp_session(self.page)
            self._cdp.send("DOM.enable")
            self._cdp.send("CSS.enable")
        root = self._cdp.send("DOM.getDocument", {"depth": 0})["root"]["nodeId"]
        node = self._cdp.send("DOM.querySelector", {"nodeId": root, "selector": "#__d2p_font_probe"})["nodeId"]
        return self._cdp.send("CSS.getPlatformFontsForNode", {"nodeId": node})["fonts"]

    def show(self, i: int, use_keyboard: bool) -> None:
        if use_keyboard and i > 0:
            # The deck's own controller runs JS-driven entrance logic that class-toggling would skip.
            self.page.keyboard.press("ArrowRight")
            for _ in range(12):
                if self.page.evaluate("() => __d2p.shownIndex()") == i:
                    break
                self.page.wait_for_timeout(40)
        self.page.evaluate("i => __d2p.goto(i)", i)
        self.page.evaluate("""async i => {
            await document.fonts.ready;
            const imgs = [...__d2p.slides[i].querySelectorAll('img')];
            await Promise.all(imgs.map(m => m.decode ? m.decode().catch(() => {}) : null));
        }""", i)
        last, deadline = None, time.time() + 1.5
        while time.time() < deadline:
            self.page.evaluate("i => __d2p.settle(i)", i)
            stamp = self.page.evaluate("i => __d2p.textStamp(i)", i)
            if stamp == last:
                break
            last = stamp
            self.page.wait_for_timeout(80)
        self.page.evaluate("() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)))")

    def screenshot(self) -> Image.Image:
        png = self.page.screenshot(type="png", timeout=60_000)
        return Image.open(io.BytesIO(png)).convert("RGB")

    def extract(self, i: int) -> dict:
        self.page.evaluate("i => { __d2p.prepare(i); __d2p.rememberHidden(i); }", i)
        return self.page.evaluate("i => __d2p.extract(i)", i)

    def capture(self, specs: list[dict], clip: dict, scale: str = "device") -> Image.Image:
        """Pixels of just the listed elements on transparency. `specs`: {id, mode}, mode own|tree."""
        self.page.evaluate("s => __d2p.showOnly(s)", specs)
        x0, y0 = max(0, clip["x"]), max(0, clip["y"])
        x1 = min(self.viewport[0], clip["x"] + clip["w"])
        y1 = min(self.viewport[1], clip["y"] + clip["h"])
        area = {"x": x0, "y": y0, "width": max(1, x1 - x0), "height": max(1, y1 - y0)}
        png = self.page.screenshot(type="png", omit_background=True, clip=area, scale=scale, timeout=60_000)
        self.page.evaluate("() => __d2p.clearShow()")
        return Image.open(io.BytesIO(png)).convert("RGBA")
