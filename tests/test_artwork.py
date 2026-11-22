import hashlib
from pathlib import Path
import tempfile
import unittest

from PIL import Image

from paperflow.artwork import stage_artwork


_EXAMPLE = Path(__file__).resolve().parents[1] / "examples/computers-and-structures"


def _placement(*captions):
    return {"placement_verified": True, "copies": [
        {"kind": "figure", "caption": caption} for caption in captions
    ]}


class ArtworkTests(unittest.TestCase):
    def test_vector_fixture_is_staged_with_original_pdf_bytes(self):
        source = _EXAMPLE / "response.pdf"
        with tempfile.TemporaryDirectory() as temp:
            report = stage_artwork(
                _placement("Figure 1 Synthetic response"), {"figure_sources": {"1": "response.pdf"}},
                _EXAMPLE, Path(temp) / "artwork",
            )
            result = report["figures"][0]
            self.assertTrue(report["all_figures_ready"])
            self.assertFalse(report["visual_match_verified"])
            self.assertEqual(Path(result["output"]).name, "Figure_1.pdf")
            self.assertEqual(Path(result["output"]).read_bytes(), source.read_bytes())
            self.assertEqual(result["sha256"], hashlib.sha256(source.read_bytes()).hexdigest())
            self.assertEqual(result["checks"]["painted_raster_images"], 0)
            self.assertGreater(result["checks"]["vector_paths"], 0)
            self.assertTrue(all(font["embedded"] for font in result["checks"]["fonts"]))

    def test_raster_requirements_preserve_bytes_even_when_quality_fails(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            for name, width, dpi in (("photo.png", 1063, 300), ("line.png", 3543, 300)):
                Image.new("RGB", (width, 20)).save(root / name, dpi=(dpi, dpi))
            metadata = {"figure_sources": {
                "1": {"path": "photo.png", "kind": "photo", "layout": "single_column"},
                "2": {"path": "line.png", "kind": "line", "layout": "single_column"},
            }}
            report = stage_artwork(_placement("Fig. 1 Photograph", "Figure 2 Line drawing"), metadata, root, root / "out")
            photo, line = report["figures"]
            self.assertEqual(photo["status"], "verified")
            self.assertEqual(line["status"], "failed")
            self.assertIn("Declared resolution falls below 1000 dpi", line["issues"])
            self.assertEqual(Path(line["output"]).read_bytes(), (root / "line.png").read_bytes())
            self.assertTrue(report["all_figures_staged"])
            self.assertFalse(report["all_figures_ready"])

    def test_missing_artwork_stays_unresolved(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            report = stage_artwork(
                _placement("Figure 1 Undeclared", "Figure 2 Missing file"),
                {"figure_sources": {"2": "missing.pdf"}},
                root, root / "out",
            )
            self.assertEqual([record["status"] for record in report["figures"]], ["missing", "missing"])
            self.assertFalse(report["all_figures_ready"])
            self.assertFalse(report["all_figures_staged"])
            self.assertEqual(list((root / "out").iterdir()), [])

    def test_chapter_caption_does_not_stage_artwork_for_its_prefix_only(self):
        with tempfile.TemporaryDirectory() as temp:
            report = stage_artwork(
                _placement("Figure \u200e1-1 Response"), {"figure_sources": {"1": "response.pdf"}},
                _EXAMPLE, Path(temp) / "artwork",
            )
            self.assertEqual(report["figures"][0]["status"], "manual_review")
            self.assertEqual(report["unused_declarations"], ["1"])
            self.assertFalse(report["all_figures_staged"])

    def test_unsafe_paths_are_rejected_without_copying(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            (root / "source").mkdir()
            (root / "private.pdf").write_bytes(b"Outside source directory")
            report = stage_artwork(
                _placement("Figure 1 Illustration"),
                {"figure_sources": {"1": "../private.pdf"}},
                root / "source", root / "out",
            )
            self.assertEqual(report["figures"][0]["status"], "failed")
            self.assertFalse(report["all_figures_staged"])
            self.assertEqual(list((root / "out").iterdir()), [])


if __name__ == "__main__":
    unittest.main()
