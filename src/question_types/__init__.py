"""Question type processors for different survey question types."""

from .single_select import cut_single_select
from .single_select_with_other import cut_single_select_with_other

__all__ = [
    'cut_single_select',
    'cut_single_select_with_other',
]
