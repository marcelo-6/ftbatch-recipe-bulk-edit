"""
XMLWriter: serialize updated RecipeTree models back to XML files.
"""

import os
from datetime import datetime
from collections.abc import Callable


class XMLWriter:
    """
    Write one or more RecipeTree instances to timestamped output folder.
    """

    def write(
        self,
        trees: list,
        base_dir: str = None,
        progress_cb: Callable[[str, dict], None] | None = None,
    ) -> str:
        """
        Serialize all modified RecipeTree objects back to XML files in a timestamped subfolder.

        This method determines an output root (either `base_dir` or a `converted-outputs` folder
        alongside the first XML), creates a timestamped directory (YYYY-MM-DD-HHMM), and writes
        each RecipeTree's updated `.tree` to a file of the same name.  Before writing, it calls
        each node's `reorder_children()` to ensure canonical tag order.  The XML is emitted with
        `utf-8` encoding, an XML declaration, and pretty-printed indentation.  Finally, it returns
        the full path of the created output directory for downstream use or user notification.

        Args:
            trees: List of RecipeTree with modifications applied.
            base_dir: Optional root dir for outputs; defaults to parent of first tree's path.

        Returns:
            output_dir: The full path of the created output folder.
        """
        def _emit(event: str, **payload) -> None:
            if progress_cb is not None:
                progress_cb(event, payload)

        total = len(trees)
        _emit("start", total=total)

        first = trees[0]
        root_dir = base_dir or os.path.join(
            os.path.dirname(first.filepath), "converted-outputs"
        )
        stamp = datetime.now().strftime("%Y-%m-%d-%H%M")
        out_dir = os.path.join(root_dir, stamp)
        os.makedirs(out_dir, exist_ok=True)

        for index, t in enumerate(trees, start=1):
            for node in t.parameters + t.formula_values:
                node.reorder_children()
            fname = os.path.basename(t.filepath)
            out_path = os.path.join(out_dir, fname)
            t.tree.write(
                out_path, encoding="utf-8", xml_declaration=True, pretty_print=True
            )
            _emit(
                "file_written",
                index=index,
                total=total,
                filename=fname,
                output_path=out_path,
            )

        _emit("finished", total=total, output_dir=out_dir)
        return out_dir
