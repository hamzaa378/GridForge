import unittest
from pathlib import Path
ROOT=Path(__file__).resolve().parent.parent
class MacroShortcutChecks(unittest.TestCase):
    def test_shortcut_points_to_macro_workbench(self):
        installer=(ROOT/'desktop/installer/GridForge.iss').read_text(encoding='utf-8')
        self.assertIn('GridForge Macro Studio', installer)
        self.assertIn('--macro-workbench', installer)
        desktop=(ROOT/'desktop/gridforge_desktop.py').read_text(encoding='utf-8')
        self.assertIn('--macro-workbench', desktop)
        self.assertIn('?macro-workbench=1', desktop)
    def test_macro_assets_bundled(self):
        main=(ROOT/'dist/main.js').read_text(encoding='utf-8')
        self.assertIn('macro-params', main)
        self.assertIn('removeDuplicates', (ROOT/'dist/engine/macros.js').read_text(encoding='utf-8'))
if __name__=='__main__':unittest.main()
