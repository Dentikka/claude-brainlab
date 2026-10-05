"""Whole-slide builders on top of `Deck`.

Geometry is written once in a 13.333 x 7.5 in design grid and scaled to the theme's slide
width (every supported format is 16:9), so the same builders serve the lab template
(10 x 5.625) and a wide co-author template (13.333 x 7.5). Font sizes that are not part of
the theme scale with `font_k` (1.0 on the wide grid).
"""
from __future__ import annotations

from typing import Optional, Sequence, Tuple

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

from .core import Deck

__all__ = ["Builders"]

GRID_W = 13.333


class Builders(Deck):
    """Slide builders; subclasses (a style) may override content/divider/recap/closing."""

    font_k: float = 1.0

    # ------------------------------------------------------------- scaling
    @property
    def k(self) -> float:
        return self.theme.width / GRID_W

    def u(self, *v: float):
        r = tuple(x * self.k for x in v)
        return r[0] if len(r) == 1 else r

    def fs(self, pt: float) -> int:
        return max(8, round(pt * self.font_k))

    # ---------------------------------------------------------- base slides
    def content(self, title: str, source: Optional[str] = None, url: Optional[str] = None):
        """Light slide with a title (and a citation line when given)."""
        t = self.theme
        s = self.new_slide()
        self.text(s, *t.title_box, title, size=t.title_size, color=t.c("title"), font=t.title_font)
        if source:
            self.cite(s, source, url)
        return s

    def divider(self, kicker: str, title: str, sub: str = ""):
        t = self.theme
        s = self.new_slide(dark=True)
        self.text(s, *self.u(0.94, 2.74, 11.44, 0.30), kicker, size=self.fs(14), color=t.c("accent2"),
                  align=PP_ALIGN.CENTER, spc=200)
        self.text(s, *self.u(0.94, 3.26, 11.44, 0.90), title, size=self.fs(54), color=t.c("gold"),
                  font=t.title_font, bold=True, align=PP_ALIGN.CENTER)
        if sub:
            self.text(s, *self.u(1.66, 4.39, 10.01, 0.80), sub, size=self.fs(20), color=t.c("on_dark"),
                      align=PP_ALIGN.CENTER)
        return s

    def takeaway(self, s, label: str, body: str, y: float = 5.35, h: float = 0.62):
        """Gold strip with a bold label and one sentence; y, h in grid inches."""
        t = self.theme
        self.card(s, *self.u(0.89, y, 11.56, h), t.c("gold"))
        self.text(s, *self.u(1.11, y, 11.1, h), [[(label + "  ", True), (body, False)]], size=self.fs(14),
                  anchor=MSO_ANCHOR.MIDDLE)

    def recap(self, title: str, items: Sequence[Tuple[str, str]], footer: str = ""):
        """Block takeaways: up to three cards (kicker, sentence) on a gold slide."""
        t = self.theme
        s = self.new_slide(bg=t.c("gold"))
        self.text(s, *t.title_box, title, size=t.title_size, color=t.c("title_dark"), font=t.title_font)
        for i, (kick, body) in enumerate(items):
            y = 1.51 + 1.2 * i
            self.card(s, *self.u(0.89, y, 11.56, 1.0), t.c("bg"), radius=0.08)
            self.text(s, *self.u(1.17, y + 0.20, 11.0, 0.26), kick.upper(), size=self.fs(12),
                      color=t.c("accent"), bold=True, spc=150)
            self.text(s, *self.u(1.17, y + 0.48, 11.0, 0.40), body, size=self.fs(18))
        if footer:
            self.text(s, *self.u(0.89, 6.12, 11.90, 0.30), footer, size=self.fs(14), color=t.c("title_dark"))
        return s

    # ------------------------------------------------------- paper slides
    def paper(self, title, lead, items, fig, caption, tab, take, source, url):
        """Lead + three bullets on the left, figure on the right, takeaway strip below."""
        t = self.theme
        s = self.content(title, source, url)
        self.text(s, *self.u(0.89, 1.52, 6.2, 0.70), lead, size=t.lead_size, color=t.c("title"))
        self.bullets(s, *self.u(0.89, 2.40, 6.2, 2.8), items, size=t.body_size)
        self.picture(s, fig, *self.u(7.35, 1.52, 5.1, 3.25))
        self.text(s, *self.u(7.35, 4.85, 5.1, 0.30), caption, size=t.small_size, color=t.c("muted"),
                  align=PP_ALIGN.CENTER)
        self.takeaway(s, tab, take)
        return s

    def two_col(self, title, head1, body1, head2, items, fig, caption, source, url):
        """Two text blocks on the left (idea, what matters), figure on the right."""
        t = self.theme
        s = self.content(title, source, url)
        self.text(s, *self.u(0.89, 1.52, 6.2, 0.33), head1, size=t.lead_size, color=t.c("title"))
        self.text(s, *self.u(0.89, 1.90, 6.2, 0.80), body1, size=t.body_size)
        self.text(s, *self.u(0.89, 2.90, 6.2, 0.33), head2, size=t.lead_size, color=t.c("title"))
        self.bullets(s, *self.u(0.89, 3.30, 6.2, 2.9), items, size=t.body_size)
        self.picture(s, fig, *self.u(7.35, 1.52, 5.1, 3.9))
        self.text(s, *self.u(7.35, 5.50, 5.1, 0.30), caption, size=t.small_size, color=t.c("muted"),
                  align=PP_ALIGN.CENTER)
        return s

    def stats(self, title, cards: Sequence[Tuple[str, str, str]], tab, take, source=None, url=None):
        """Three big numbers (value, label, sub) and a takeaway strip."""
        t = self.theme
        s = self.content(title, source, url)
        for i, (value, label, sub) in enumerate(cards):
            x = 0.89 + 3.93 * i
            self.card(s, *self.u(x, 1.54, 3.70, 1.75), t.c("card_alt") if i == 1 else t.c("card"))
            self.text(s, *self.u(x + 0.22, 1.72, 3.3, 0.60), value, size=self.fs(36), bold=True,
                      color=t.c("accent") if i == 1 else t.c("title"))
            self.text(s, *self.u(x + 0.22, 2.40, 3.3, 0.30), label, size=t.body_size)
            self.text(s, *self.u(x + 0.22, 2.75, 3.3, 0.40), sub, size=t.small_size, color=t.c("muted"))
        self.takeaway(s, tab, take, y=3.65, h=0.9)
        return s

    def cards3(self, title, lead, cards: Sequence[Tuple[str, str, str, str]], foot="", source=None, url=None):
        """Lead line + three cards (kicker, heading, body, footnote) + a closing line."""
        t = self.theme
        s = self.content(title, source, url)
        self.text(s, *self.u(0.89, 1.52, 11.56, 0.40), lead, size=t.lead_size, color=t.c("title"))
        for i, (kick, head, body, sub) in enumerate(cards):
            x = 0.89 + 3.93 * i
            self.card(s, *self.u(x, 2.15, 3.70, 2.65), t.c("card_alt") if i == 1 else t.c("card"))
            self.text(s, *self.u(x + 0.22, 2.33, 3.3, 0.26), kick.upper(), size=self.fs(12), color=t.c("accent"),
                      bold=True, spc=150)
            self.text(s, *self.u(x + 0.22, 2.65, 3.3, 0.40), head, size=t.lead_size, color=t.c("title"))
            self.text(s, *self.u(x + 0.22, 3.10, 3.3, 1.0), body, size=t.body_size)
            self.text(s, *self.u(x + 0.22, 4.25, 3.3, 0.45), sub, size=t.small_size, color=t.c("muted"))
        if foot:
            self.text(s, *self.u(0.89, 5.10, 11.56, 0.30), foot, size=self.fs(14), color=t.c("title"))
        return s

    def measurement(self, title, lead, rows, widths, note, tab, take, provenance):
        """Our own measurement: one table, one muted line on how numbers were obtained
        (measured vs derived), a takeaway strip and a provenance line instead of a citation."""
        t = self.theme
        s = self.content(title, provenance)
        self.text(s, *self.u(0.89, 1.52, 11.56, 0.40), lead, size=t.lead_size, color=t.c("title"))
        self.table(s, rows, [w * self.k for w in widths], x=self.u(0.90), y=self.u(2.20), row_h=self.u(0.55),
                   size=self.fs(15))
        self.text(s, *self.u(0.89, 5.00, 11.56, 0.30), note, size=t.small_size, color=t.c("muted"))
        self.takeaway(s, tab, take, y=5.45, h=0.95)
        return s

    def closing(self, title: str = "Thank you for your attention!", picture: Optional[str] = None):
        t = self.theme
        s = self.content(title)
        if picture:
            self.picture(s, picture, *self.u(3.47, 1.62, 6.4, 5.0))
        return s
