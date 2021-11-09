using System.Collections.Generic;
using System.Drawing;
using DevExpress.XtraBars;
using DevExpress.XtraBars.Ribbon;
using DevExpress.XtraSpellChecker;
using DevExpress.XtraSpellChecker.Native;
using DevExpress.XtraSpreadsheet.Internal;

namespace SpreadsheetSpellchecking {
    public partial class Form1 : RibbonForm {
        SpellChecker spellChecker = new SpellChecker();
        PopupMenu popupMenu = new PopupMenu();

        public Form1() {
            InitializeComponent();
            SpellCheckTextControllersManager.Default.RegisterClass(typeof(TextBoxWithTransparency), typeof(SimpleTextBoxTextController));
            SpellCheckTextBoxBaseFinderManager.Default.RegisterClass(typeof(TextBoxWithTransparency), typeof(TextBoxFinder));

            spellChecker.SpellCheckMode = SpellCheckMode.AsYouType;

            popupMenu.Manager = BarManager;
            BarManager.QueryShowPopupMenu += OnBarManager_QueryShowPopupMenu;

            spreadsheetControl.CellEditorOpened += OnSpreadsheetControl_CellEditorOpened;
        }

        BarManager BarManager => ribbonControl1.Manager;

        void OnSpreadsheetControl_CellEditorOpened(object sender, DevExpress.XtraSpreadsheet.CellEditorOpenedEventArgs e) {
            if (!e.IsCustom) {
                BarManager.SetPopupContextMenu(e.Editor, popupMenu);
                spellChecker.Check(e.Editor);
            }
        }

        void OnBarManager_QueryShowPopupMenu(object sender, QueryShowPopupMenuEventArgs e) {
            Point position = e.Control.PointToClient(e.Position);
            DevExpress.XtraSpellChecker.Rules.SpellCheckErrorBase error = spellChecker.CalcError(position);
            List<SpellCheckerCommand> commands = spellChecker.GetCommandsByError(error);
            if (commands == null) {
                e.Cancel = true;
                return;
            }

            BarItemLinkCollection itemLinks = popupMenu.ItemLinks;
            popupMenu.BeginUpdate();
            try {
                itemLinks.Clear();
                foreach (SpellCheckerCommand command in commands) {
                    BarButtonItem item = new BarButtonItem(BarManager, command.Caption);
                    item.Enabled = command.Enabled;
                    item.Tag = command;
                    item.ItemClick += OnPopupMenu_ItemClick;
                    itemLinks.Add(item);
                }
                BarButtonItem itemShowSpellingForm = new BarButtonItem(BarManager, "Show Spelling Form");
                itemShowSpellingForm.ItemClick += OnPopupMenuShowSpellingForm_ItemClick;
                itemLinks.Add(itemShowSpellingForm);
            } finally {
                popupMenu.EndUpdate();
            }
        }
        void OnPopupMenu_ItemClick(object sender, ItemClickEventArgs e) {
            var checkerCommand = e.Item.Tag as SpellCheckerCommand;
            checkerCommand?.DoCommand();
        }
        void OnPopupMenuShowSpellingForm_ItemClick(object sender, ItemClickEventArgs e) {
            spellChecker.Check(this.ActiveControl);
        }

        private void Form1_Load(object sender, System.EventArgs e)
        {
            var activeSheet = spreadsheetControl.ActiveWorksheet;
            activeSheet.Columns[1].WidthInPixels = 150;
            spreadsheetControl.SelectedCell = activeSheet["B2"];
            spreadsheetControl.SelectedCell.Value = "Missspelled wods";
        }
    }
}