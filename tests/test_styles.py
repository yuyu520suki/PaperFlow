"""Ensure native CSL installation remains opt-in and dependency bounded."""

import hashlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from paperflow.styles import install_styles


def write_style(directory, name, parent=None):
    ident = f"http://www.zotero.org/styles/{name}"
    parent_link = f'<link rel="independent-parent" href="{parent}"/>' if parent else ""
    data = f'<style xmlns="http://purl.org/net/xbiblio/csl" version="1.0"><info><title>{name}</title><id>{ident}</id>{parent_link}</info></style>'.encode()
    path = directory / f"{name}.csl"
    path.write_bytes(data)
    return path, hashlib.sha256(data).hexdigest()


class StyleInstallerTests(unittest.TestCase):
    def test_default_dry_run_does_not_contact_zotero_and_sorts_dependencies(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            styles = root / "installed"
            styles.mkdir()
            parent, parent_hash = write_style(root, "parent")
            dependent, dependent_hash = write_style(root, "dependent", "http://www.zotero.org/styles/parent")
            with patch("paperflow.styles.urllib.request.build_opener") as opener:
                result = install_styles([dependent, parent], expected_sha256={parent.name: parent_hash, dependent.name: dependent_hash}, styles_dir=styles)
                opener.assert_not_called()
            self.assertFalse(result["writes_performed"])
            self.assertEqual([s["filename"] for s in result["plan"]], ["parent.csl", "dependent.csl"])
            self.assertEqual(list(styles.iterdir()), [])

    def test_wrong_digest_and_missing_dependency_fail_before_network(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            styles = root / "installed"
            styles.mkdir()
            dependent, digest = write_style(root, "dependent", "http://www.zotero.org/styles/missing")
            with patch("paperflow.styles.urllib.request.build_opener") as opener:
                with self.assertRaises(ValueError):
                    install_styles([dependent], expected_sha256={dependent.name: "0" * 64}, styles_dir=styles, install=True)
                with self.assertRaises(ValueError):
                    install_styles([dependent], expected_sha256={dependent.name: digest}, styles_dir=styles, install=True)
                opener.assert_not_called()

    def test_existing_different_style_is_not_overwritten(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            styles = root / "installed"
            styles.mkdir()
            original, _ = write_style(styles, "same")
            original_data = original.read_bytes()
            changed, _ = write_style(root, "same")
            changed.write_bytes(changed.read_bytes().replace(b"<title>same</title>", b"<title>changed</title>"))
            digest = hashlib.sha256(changed.read_bytes()).hexdigest()
            with self.assertRaises(ValueError):
                install_styles([changed], expected_sha256={changed.name: digest}, styles_dir=styles, install=True)
            self.assertEqual(original.read_bytes(), original_data)


if __name__ == "__main__":
    unittest.main()
