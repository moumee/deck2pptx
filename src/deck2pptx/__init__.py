"""Convert fixed-stage HTML slide decks into editable PowerPoint files."""
from .convert import convert
from .update import inspect, update

__version__ = "0.2.0"
__all__ = ["convert", "inspect", "update"]
