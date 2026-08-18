from __future__ import annotations

import unittest
from pathlib import Path


REPOSITORY_ROOT = Path(__file__).resolve().parents[1]
WINDOWS_DIR = REPOSITORY_ROOT / "windows"
LAUNCHER = WINDOWS_DIR / "Graph-App-Role-Manager.ps1"
MODULE = WINDOWS_DIR / "Graph-App-Role-Manager.psm1"

LAUNCHER_TEXT = LAUNCHER.read_text(encoding="utf-8")
MODULE_TEXT = MODULE.read_text(encoding="utf-8")


class WindowsModuleArchitectureTests(unittest.TestCase):
    def test_launcher_and_module_files_exist(self) -> None:
        self.assertTrue(LAUNCHER.is_file(), "The native Windows launcher must remain available.")
        self.assertTrue(MODULE.is_file(), "The native Windows module must exist.")

    def test_launcher_imports_module_from_psscriptroot(self) -> None:
        self.assertIn("Join-Path $PSScriptRoot 'Graph-App-Role-Manager.psm1'", LAUNCHER_TEXT)
        self.assertIn("Import-Module $modulePath -Force", LAUNCHER_TEXT)

    def test_launcher_invokes_public_module_entry_point(self) -> None:
        self.assertIn("Start-GraphAppRoleManager", LAUNCHER_TEXT)
        self.assertIn("-LogLevel $LogLevel", LAUNCHER_TEXT)

    def test_module_exports_start_graph_app_role_manager(self) -> None:
        self.assertIn("function Start-GraphAppRoleManager", MODULE_TEXT)
        self.assertIn(
            "Export-ModuleMember -Function Start-GraphAppRoleManager",
            MODULE_TEXT,
        )

    def test_log_level_uses_validated_set_and_defaults_to_info(self) -> None:
        for level in ("INFO", "SUCCESS", "WARNING", "ERROR"):
            self.assertIn(level, MODULE_TEXT)

        self.assertRegex(
            MODULE_TEXT,
            r"\[ValidateSet\('INFO', 'SUCCESS', 'WARNING', 'ERROR'\)\][\s\S]*?\[string\]\$LogLevel = 'INFO'",
        )
        self.assertIn("$script:LogLevel = 'INFO'", MODULE_TEXT)

    def test_log_level_is_not_stored_in_global_scope(self) -> None:
        self.assertNotIn("$Global:Globalloglevel", MODULE_TEXT)
        self.assertNotIn("$Global:LogLevel", MODULE_TEXT)
        self.assertIn("$script:LogLevel", MODULE_TEXT)

    def test_safety_sensitive_behaviors_remain_present(self) -> None:
        safety_patterns = (
            r"\$missingModules\s*=\s*@\(\s*foreach",
            r"Load-GraphApplicationRoles\s+Load-TargetServicePrincipals",
            r"Confirm permission assignment",
            r"Confirm permission removal",
            r"Disconnect-MgGraph",
        )
        for pattern in safety_patterns:
            with self.subTest(pattern=pattern):
                self.assertRegex(MODULE_TEXT, pattern)

    def test_launcher_forwards_log_level_parameter(self) -> None:
        self.assertIn(
            "[ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR')]",
            LAUNCHER_TEXT,
        )
        self.assertIn("[string]$LogLevel = 'INFO'", LAUNCHER_TEXT)

    def test_module_state_is_initialized_before_reset_reads_controls(self) -> None:
        self.assertIn("function Initialize-GraphAppRoleManagerModuleState", MODULE_TEXT)
        self.assertIn("$script:form = $null", MODULE_TEXT)
        self.assertRegex(
            MODULE_TEXT,
            r"function Reset-GraphAppRoleManagerState \{[\s\S]*?\$null -ne \$script:form[\s\S]*?Initialize-GraphAppRoleManagerModuleState",
        )
        self.assertRegex(
            MODULE_TEXT,
            r"Initialize-GraphAppRoleManagerModuleState\s*\n\s*\n# ---------------------------------------------------------------------------\n# HELPERS",
        )
        reset_start = MODULE_TEXT.index("function Reset-GraphAppRoleManagerState")
        form_init = MODULE_TEXT.index("$script:form = $null")
        self.assertLess(
            form_init,
            reset_start,
            "Module GUI state must be initialized before Reset reads $script:form.",
        )


if __name__ == "__main__":
    unittest.main()
