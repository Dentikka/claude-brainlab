"""deckkit: build talk and lecture decks as code (python-pptx).

    from deckkit import LabDeck                         # lab house style (default)
    from deckkit import Builders, Theme                 # a co-author's or any custom style
    from deckkit.course import open_course, CourseDeck  # LLM-course template
    from deckkit import checks                          # preflight checks on any .pptx
"""
from . import checks
from .builders import Builders
from .core import Deck, Theme, arxiv_abs, arxiv_pdf, nav_pills, rgb
from .lab import LAB, LabDeck

__all__ = ["Builders", "Deck", "Theme", "LAB", "LabDeck", "arxiv_abs", "arxiv_pdf", "checks", "nav_pills", "rgb"]
