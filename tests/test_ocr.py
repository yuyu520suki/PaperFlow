import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from paperflow.ocr import OcrError, recognize_image


class OcrContractTests(unittest.TestCase):
    def test_ocr_output_cannot_claim_formula_or_table_reconstruction(self):
        with tempfile.TemporaryDirectory() as temp:
            image, output = Path(temp)/"input.png", Path(temp)/"output.json"
            image.write_bytes(b"fixture")
            payload={"engine":"Windows.Media.Ocr","text":"x = 1","lines":[{"text":"x = 1","words":[]}],"requires_review":False}
            with patch("paperflow.ocr._invoke",return_value=payload):
                report=recognize_image(image,output,language="en-US")
            self.assertTrue(report["requires_review"])
            self.assertFalse(report["equations_reconstructed"])
            self.assertFalse(report["table_structure_reconstructed"])
            self.assertEqual(json.loads(output.read_text(encoding="utf-8"))["text"],"x = 1")

    def test_existing_output_is_protected_before_ocr_is_started(self):
        with tempfile.TemporaryDirectory() as temp:
            image, output = Path(temp)/"input.png", Path(temp)/"output.json"
            image.write_bytes(b"fixture"); output.write_text("existing",encoding="utf-8")
            with patch("paperflow.ocr._invoke") as invoke:
                with self.assertRaises(FileExistsError): recognize_image(image,output)
                invoke.assert_not_called()
            self.assertEqual(output.read_text(encoding="utf-8"),"existing")

    def test_non_image_or_invalid_language_does_not_start_ocr(self):
        with tempfile.TemporaryDirectory() as temp:
            source=Path(temp)/"input.pdf"; source.write_bytes(b"fixture")
            with patch("paperflow.ocr._invoke") as invoke:
                with self.assertRaises(ValueError): recognize_image(source)
                image=Path(temp)/"input.png"; image.write_bytes(b"fixture")
                with self.assertRaises(ValueError): recognize_image(image,language="en-US; bad")
                invoke.assert_not_called()

    def test_malformed_engine_output_is_not_published(self):
        with tempfile.TemporaryDirectory() as temp:
            image, output = Path(temp)/"input.png", Path(temp)/"output.json"
            image.write_bytes(b"fixture")
            with patch("paperflow.ocr._invoke",return_value={"engine":"other","text":"text","lines":[]}):
                with self.assertRaises(OcrError): recognize_image(image,output)
            self.assertFalse(output.exists())


if __name__=="__main__": unittest.main()
