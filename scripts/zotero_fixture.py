"""Resolve public smoke-test citations in the current local Zotero library."""
from functools import lru_cache

from paperflow.doi import resolve_dois


FIRST_DOI = "10.6052/j.issn.1000-4750.2020.11.0786"
SECOND_DOI = "10.1016/j.compstruc.2022.106899"


@lru_cache(maxsize=2)
def fixture_key(doi):
    matches = resolve_dois([doi])[doi]
    if len(matches) != 1:
        raise ValueError(f"This native smoke test needs one local Zotero record with DOI {doi}.")
    return matches[0]["key"]
