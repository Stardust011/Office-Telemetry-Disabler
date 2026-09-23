import re
import unittest
from pathlib import Path

REPO = Path('/home/runner/work/Office-Telemetry-Disabler/Office-Telemetry-Disabler')
SCRIPT = (REPO / 'script' / 'office_privacy_telemetry_disabler.ps1').read_text(encoding='utf-8')
REG = (REPO / 'OC2R_DisableTelemetry.reg').read_text(encoding='utf-16le')


def normalize_path(path: str) -> str:
    p = path.replace('HKEY_CURRENT_USER', 'HKCU')
    p = p.replace('HKEY_LOCAL_MACHINE', 'HKLM')
    p = p.replace('\\', '\\')
    return p.lower()


class ValidateOffice16Profile(unittest.TestCase):
    def test_only_office_16_target(self):
        self.assertIn("$OfficeVersion = '16.0'", SCRIPT)
        self.assertNotIn('14.0', SCRIPT)
        self.assertNotIn('15.0', SCRIPT)

    def test_update_disable_is_optional_and_default_off(self):
        self.assertRegex(SCRIPT, r"disableUpdatesAnswer\s*=\s*Read-Host")
        self.assertRegex(SCRIPT, r"\$disableUpdates\s*=\s*\$disableUpdatesAnswer\s*-match\s*'\^\[Yy\]\$'")
        self.assertIn('if ($disableUpdates)', SCRIPT)

    def test_baseline_reg_entries_exist_in_script(self):
        key = None
        reg_pairs = []
        for raw in REG.splitlines():
            line = raw.strip()
            if not line:
                continue
            if line.startswith('[') and line.endswith(']'):
                key = line[1:-1]
                continue
            if line.startswith('"') and '=' in line and key:
                name = line.split('"=')[0].strip('"')
                reg_pairs.append((normalize_path(key), name.lower()))

        # Build lookup from script hashtable entries
        script_pairs = set()
        for m in re.finditer(r"Path\s*=\s*'([^']+)'\s*;\s*Name\s*=\s*'([^']+)'", SCRIPT):
            script_path = m.group(1).replace('$OfficeVersion', '16.0').replace(':', '')
            script_pairs.add((normalize_path(script_path), m.group(2).lower()))
        for m in re.finditer(r'Path\s*=\s*"([^"]+)"\s*;\s*Name\s*=\s*\'([^\']+)\'', SCRIPT):
            script_path = m.group(1).replace('$OfficeVersion', '16.0').replace(':', '')
            script_pairs.add((normalize_path(script_path), m.group(2).lower()))

        missing = [pair for pair in reg_pairs if pair not in script_pairs]
        self.assertFalse(missing, f'Missing reg baseline entries in script: {missing[:5]}')


if __name__ == '__main__':
    unittest.main()
