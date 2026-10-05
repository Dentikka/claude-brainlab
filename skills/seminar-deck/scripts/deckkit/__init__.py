"""deckkit: build talk and lecture decks as code (python-pptx).

    from deckkit import LabDeck                         # lab house style (default)
    from deckkit import Builders, Theme                 # a co-author's or any custom style
    from deckkit.course import open_course, CourseDeck  # LLM-course template
"""
from .builders import Builders
from .core import Deck, Theme, arxiv_abs, arxiv_pdf, rgb
from .lab import LAB, LabDeck

__all__ = ["Builders", "Deck", "Theme", "LAB", "LabDeck", "arxiv_abs", "arxiv_pdf", "rgb"]
