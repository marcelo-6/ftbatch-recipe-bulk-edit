"""
XMLParser: recursively load parent and child XMLs into RecipeTree instances.
"""

import os
import logging
from collections import Counter
from collections.abc import Callable
from core.xml_model import RecipeTree
from core.base import NSMAP


class XMLParser:
    """
    Parses a .pxml/.uxml tree with recursive child loading.
    """

    def parse(
        self,
        parent_path: str,
        progress_cb: Callable[[str, dict], None] | None = None,
    ) -> list:
        """
        Recursively load a parent XML recipe and all of its child recipe files into RecipeTree objects.

        This method begins by loading the parent `.pxml` (or `.uxml`) file, creating a RecipeTree
        instance that captures its parameters and formula values.  It then inspects every
        `<StepRecipeID>` element to discover child files (e.g. `.uxml` or `.oxml`) in the same directory,
        loading each of those exactly once.  Per-file diagnostics are logged at DEBUG, while warnings are
        summarized if child files are missing.  Finally, it returns a list of all distinct RecipeTree
        objects (parent plus children), preserving the original loading order.
        """
        loaded = {}
        discovered_paths = {os.path.abspath(parent_path)}
        missing_children: Counter[str] = Counter()
        log = logging.getLogger(__name__)

        def _emit(event: str, **payload) -> None:
            if progress_cb is not None:
                progress_cb(event, payload)

        def _load(path):
            abs_path = os.path.abspath(path)
            if abs_path in loaded:
                return
            discovered_paths.add(abs_path)
            log.debug("Parsing XML: %s", abs_path)
            tree = RecipeTree(abs_path)
            tree.extract_nodes()
            loaded[abs_path] = tree

            log.debug(
                f"\tLoaded {os.path.basename(abs_path)}: {len(tree.parameters)} params, {len(tree.formula_values)} formula values, Total = {len(tree.parameters) + len(tree.formula_values)}"
            )
            _emit(
                "loaded",
                path=abs_path,
                loaded=len(loaded),
                total=len(discovered_paths),
                params=len(tree.parameters),
                formula_values=len(tree.formula_values),
            )
            # determine child extension
            ext = os.path.splitext(abs_path)[1].upper()
            child_ext = {".PXML": ".UXML", ".UXML": ".OXML"}.get(ext)
            if child_ext:
                for sr in tree.tree.findall(
                    f".//{{{NSMAP[None]}}}StepRecipeID", namespaces=NSMAP
                ):
                    name = (sr.text or "").strip()
                    if not name:
                        continue
                    child = os.path.abspath(
                        os.path.join(os.path.dirname(abs_path), name + child_ext)
                    )
                    log.debug(
                        f"\tParent {os.path.basename(abs_path)} - Looking for Child XML: {child}",
                    )
                    if os.path.exists(child):
                        if child not in discovered_paths and child not in loaded:
                            discovered_paths.add(child)
                            _emit(
                                "discovered",
                                path=child,
                                loaded=len(loaded),
                                total=len(discovered_paths),
                            )
                        log.debug(
                            f"\tParent {os.path.basename(abs_path)} - Child found, Parsing Child XML: {child}"
                        )
                        _load(child)
                    else:
                        missing_children[child] += 1
                        log.debug(
                            f"\tParent {os.path.basename(abs_path)} - Child XML not found: {child}"
                        )
                        _emit(
                            "missing_child",
                            path=child,
                            parent=os.path.basename(abs_path),
                            count=missing_children[child],
                        )

        _load(parent_path)
        if missing_children:
            for child, count in sorted(
                missing_children.items(),
                key=lambda item: (-item[1], item[0]),
            ):
                if count == 1:
                    log.warning("Child XML not found: %s", child)
                else:
                    log.warning("Child XML not found (%dx): %s", count, child)
        log.info(
            "Parsed XML graph: loaded=%d files, discovered=%d references, missing-child occurrences=%d",
            len(loaded),
            len(discovered_paths),
            sum(missing_children.values()),
        )
        _emit(
            "finished",
            loaded=len(loaded),
            total=len(discovered_paths),
            missing_occurrences=sum(missing_children.values()),
        )
        return list(loaded.values())
