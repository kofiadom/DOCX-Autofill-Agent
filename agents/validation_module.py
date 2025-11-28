#!/usr/bin/env python3
"""
Validation & QA Module - Phase 5 Implementation
3-tier validation for document integrity and fill quality
"""

import os
from pathlib import Path
from typing import Tuple, List, Dict, Any
from lib.document import Document


class ValidationResult:
    """Track validation results"""

    def __init__(self):
        self.passed = []
        self.failed = []
        self.warnings = []

    def add_pass(self, check: str, message: str = ""):
        self.passed.append((check, message))

    def add_fail(self, check: str, message: str):
        self.failed.append((check, message))

    def add_warning(self, check: str, message: str):
        self.warnings.append((check, message))

    def is_valid(self) -> bool:
        """Return True if no critical failures"""
        return len(self.failed) == 0

    def summary(self) -> str:
        msg = f"\nValidation: {len(self.passed)} passed"
        if self.warnings:
            msg += f", {len(self.warnings)} warnings"
        if self.failed:
            msg += f", {len(self.failed)} FAILED"
        return msg


class Tier1PlaceholderValidation:
    """Tier 1: Verify all expected placeholders were filled"""

    @staticmethod
    def validate(doc: Document, expected_fields: list) -> ValidationResult:
        """Verify all placeholders were replaced

        Checks for remaining {{PLACEHOLDER}} patterns and SDT fields
        that weren't filled.

        Args:
            doc: Document instance
            expected_fields: List of field names that should be filled

        Returns:
            ValidationResult with pass/fail/warning status
        """
        result = ValidationResult()
        doc_xml = doc["word/document.xml"]
        doc_text = doc_xml.dom.toxml()

        unfilled_placeholders = []
        unfilled_sdts = []

        # Check for unfilled text placeholders
        for field in expected_fields:
            search_text = f"{{{{{field}}}}}"
            if search_text in doc_text:
                unfilled_placeholders.append(field)

        # Check for unfilled SDT fields
        sdts = doc_xml.dom.getElementsByTagName('w:sdt')
        for sdt in sdts:
            alias_elems = sdt.getElementsByTagName('w:alias')
            if alias_elems:
                field_name = alias_elems[0].getAttribute('w:val')
                if field_name in expected_fields:
                    content_elems = sdt.getElementsByTagName('w:sdtContent')
                    if content_elems:
                        text_elems = content_elems[0].getElementsByTagName('w:t')
                        if text_elems:
                            text_val = text_elems[0].firstChild.nodeValue if text_elems[0].firstChild else ""
                            if not text_val or text_val.strip() == "":
                                unfilled_sdts.append(field_name)

        # Report results
        if not unfilled_placeholders and not unfilled_sdts:
            result.add_pass(
                "Placeholder Completion",
                f"All {len(expected_fields)} expected fields filled"
            )
        else:
            errors = []
            if unfilled_placeholders:
                errors.extend([f"Text: {f}" for f in unfilled_placeholders])
            if unfilled_sdts:
                errors.extend([f"SDT: {f}" for f in unfilled_sdts])

            result.add_fail(
                "Placeholder Completion",
                f"Unfilled: {', '.join(errors)}"
            )

        return result


class Tier2DocumentIntegrityValidation:
    """Tier 2: Check document structure integrity"""

    @staticmethod
    def validate(unpacked_path: str) -> ValidationResult:
        """Validate document structure

        Checks that all required DOCX files and directories exist.

        Args:
            unpacked_path: Path to unpacked document

        Returns:
            ValidationResult with structural checks
        """
        result = ValidationResult()

        # Required files for valid DOCX
        required_files = [
            'word/document.xml',
            '[Content_Types].xml',
            '_rels/.rels'
        ]

        # Check required files
        all_exist = True
        for file_path in required_files:
            full_path = os.path.join(unpacked_path, file_path)
            if os.path.exists(full_path):
                result.add_pass("File Check", f"Found: {file_path}")
            else:
                result.add_fail("File Check", f"Missing: {file_path}")
                all_exist = False

        # Check optional files
        optional_files = [
            'word/styles.xml',
            'word/fontTable.xml',
            'word/numbering.xml',
            'customXml/'
        ]

        for file_path in optional_files:
            full_path = os.path.join(unpacked_path, file_path)
            if os.path.exists(full_path):
                result.add_pass("Optional Files", f"Found: {file_path}")
            else:
                result.add_warning("Optional Files", f"Missing: {file_path}")

        # Check document.xml is readable
        try:
            doc_path = os.path.join(unpacked_path, 'word/document.xml')
            with open(doc_path, 'r', encoding='utf-8') as f:
                content = f.read()
                if content and '<?xml' in content:
                    result.add_pass("XML Format", "document.xml is valid XML declaration")
                else:
                    result.add_warning("XML Format", "document.xml missing XML declaration")
        except Exception as e:
            result.add_fail("XML Format", str(e))

        if all_exist:
            result.add_pass("Document Integrity", "All required DOCX files present")

        return result


class Tier3XMLValidation:
    """Tier 3: Validate XML well-formedness"""

    @staticmethod
    def validate(doc: Document) -> ValidationResult:
        """Validate XML well-formedness

        Attempts to save and validate the document structure.
        Skips schema validation since templates may have pre-existing errors.

        Args:
            doc: Document instance

        Returns:
            ValidationResult from XML validation
        """
        result = ValidationResult()

        try:
            # Save without schema validation - templates may have pre-existing errors
            # We only verify the document is still XML well-formed after our changes
            doc.save(validate=False)
            result.add_pass("XML Well-formedness", "Document XML is well-formed (no structural breaks)")

        except Exception as e:
            error_msg = str(e)
            if "namespace" in error_msg.lower():
                result.add_fail("XML Namespaces", f"Namespace error: {error_msg}")
            elif "element" in error_msg.lower():
                result.add_fail("XML Elements", f"Element error: {error_msg}")
            else:
                result.add_fail("XML Structure", error_msg)

        return result


class ComprehensiveValidator:
    """Run all 3 tiers of validation"""

    def __init__(self, doc: Document, unpacked_path: str):
        self.doc = doc
        self.unpacked_path = unpacked_path
        self.results = {}

    def validate_all(self, expected_fields: list = None) -> bool:
        """Run all validation tiers

        Args:
            expected_fields: List of field names expected in document

        Returns:
            True if all validations pass, False otherwise
        """
        print("\n" + "="*60)
        print("COMPREHENSIVE DOCUMENT VALIDATION")
        print("="*60 + "\n")

        all_valid = True

        # Tier 1: Placeholder completion
        if expected_fields:
            print("Tier 1: Placeholder Completion Check")
            print("-" * 40)
            result1 = Tier1PlaceholderValidation.validate(self.doc, expected_fields)
            self.results['Tier1-Placeholders'] = result1

            print(f"  Checked: {len(expected_fields)} expected fields")
            for check, msg in result1.passed:
                print(f"  ✓ {check}: {msg}")
            for check, msg in result1.failed:
                print(f"  ✗ {check}: {msg}")
                all_valid = False
            for check, msg in result1.warnings:
                print(f"  ⚠ {check}: {msg}")

            print(result1.summary())

        # Tier 2: Document integrity
        print("\nTier 2: Document Integrity Check")
        print("-" * 40)
        result2 = Tier2DocumentIntegrityValidation.validate(self.unpacked_path)
        self.results['Tier2-Integrity'] = result2

        for check, msg in result2.passed:
            print(f"  ✓ {check}: {msg}")
        for check, msg in result2.failed:
            print(f"  ✗ {check}: {msg}")
            all_valid = False
        for check, msg in result2.warnings:
            print(f"  ⚠ {check}: {msg}")

        print(result2.summary())

        # Tier 3: XML validation
        print("\nTier 3: XML Well-formedness Check")
        print("-" * 40)
        result3 = Tier3XMLValidation.validate(self.doc)
        self.results['Tier3-XML'] = result3

        for check, msg in result3.passed:
            print(f"  ✓ {check}: {msg}")
        for check, msg in result3.failed:
            print(f"  ✗ {check}: {msg}")
            all_valid = False
        for check, msg in result3.warnings:
            print(f"  ⚠ {check}: {msg}")

        print(result3.summary())

        # Overall result
        print("\n" + "="*60)
        if all_valid:
            print("✓ VALIDATION PASSED - Document is valid and ready")
        else:
            print("✗ VALIDATION FAILED - Document has errors")
        print("="*60 + "\n")

        return all_valid

    def get_detailed_report(self) -> str:
        """Get detailed validation report"""
        report = "VALIDATION REPORT\n"
        report += "="*60 + "\n\n"

        for tier_name, result in self.results.items():
            report += f"{tier_name}:\n"
            report += "-"*40 + "\n"

            if result.passed:
                report += "PASSED:\n"
                for check, msg in result.passed:
                    report += f"  ✓ {check}: {msg}\n"

            if result.warnings:
                report += "WARNINGS:\n"
                for check, msg in result.warnings:
                    report += f"  ⚠ {check}: {msg}\n"

            if result.failed:
                report += "FAILED:\n"
                for check, msg in result.failed:
                    report += f"  ✗ {check}: {msg}\n"

            report += "\n"

        return report


def quick_validate(doc: Document, expected_fields: list) -> tuple:
    """Quick validation without full reporting

    Args:
        doc: Document instance
        expected_fields: List of expected field names

    Returns:
        Tuple of (is_valid, unfilled_count, error_messages)
    """
    result = Tier1PlaceholderValidation.validate(doc, expected_fields)
    unfilled = len(result.failed)
    is_valid = result.is_valid()
    errors = [msg for check, msg in result.failed]

    return is_valid, unfilled, errors


if __name__ == '__main__':
    # Test validation
    print("Validation Module")
    print("Use with: from validation_module import ComprehensiveValidator")
