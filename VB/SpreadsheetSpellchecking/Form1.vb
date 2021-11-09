Imports System.Collections.Generic
Imports System.Drawing
Imports DevExpress.XtraBars
Imports DevExpress.XtraBars.Ribbon
Imports DevExpress.XtraSpellChecker
Imports DevExpress.XtraSpellChecker.Native
Imports DevExpress.XtraSpreadsheet.Internal

Namespace SpreadsheetSpellchecking

    Public Partial Class Form1
        Inherits RibbonForm

        Private spellChecker As SpellChecker = New SpellChecker()

        Private popupMenu As PopupMenu = New PopupMenu()

        Public Sub New()
            InitializeComponent()
            SpellCheckTextControllersManager.[Default].RegisterClass(GetType(TextBoxWithTransparency), GetType(SimpleTextBoxTextController))
            SpellCheckTextBoxBaseFinderManager.[Default].RegisterClass(GetType(TextBoxWithTransparency), GetType(TextBoxFinder))
            spellChecker.SpellCheckMode = SpellCheckMode.AsYouType
            popupMenu.Manager = BarManager
            AddHandler Me.BarManager.QueryShowPopupMenu, AddressOf OnBarManager_QueryShowPopupMenu
            AddHandler Me.spreadsheetControl.CellEditorOpened, AddressOf OnSpreadsheetControl_CellEditorOpened
        End Sub

        Private ReadOnly Property BarManager As BarManager
            Get
                Return ribbonControl1.Manager
            End Get
        End Property

        Private Sub OnSpreadsheetControl_CellEditorOpened(ByVal sender As Object, ByVal e As DevExpress.XtraSpreadsheet.CellEditorOpenedEventArgs)
            If Not e.IsCustom Then
                BarManager.SetPopupContextMenu(e.Editor, popupMenu)
                spellChecker.Check(e.Editor)
            End If
        End Sub

        Private Sub OnBarManager_QueryShowPopupMenu(ByVal sender As Object, ByVal e As QueryShowPopupMenuEventArgs)
            Dim position As Point = e.Control.PointToClient(e.Position)
            Dim [error] As DevExpress.XtraSpellChecker.Rules.SpellCheckErrorBase = spellChecker.CalcError(position)
            Dim commands As List(Of SpellCheckerCommand) = spellChecker.GetCommandsByError([error])
            If commands Is Nothing Then
                e.Cancel = True
                Return
            End If

            Dim itemLinks As BarItemLinkCollection = popupMenu.ItemLinks
            popupMenu.BeginUpdate()
            Try
                itemLinks.Clear()
                For Each command As SpellCheckerCommand In commands
                    Dim item As BarButtonItem = New BarButtonItem(BarManager, command.Caption)
                    item.Enabled = command.Enabled
                    item.Tag = command
                    AddHandler item.ItemClick, AddressOf OnPopupMenu_ItemClick
                    itemLinks.Add(item)
                Next

                Dim itemShowSpellingForm As BarButtonItem = New BarButtonItem(BarManager, "Show Spelling Form")
                AddHandler itemShowSpellingForm.ItemClick, AddressOf OnPopupMenuShowSpellingForm_ItemClick
                itemLinks.Add(itemShowSpellingForm)
            Finally
                popupMenu.EndUpdate()
            End Try
        End Sub

        Private Sub OnPopupMenu_ItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs)
            Dim checkerCommand = TryCast(e.Item.Tag, SpellCheckerCommand)
            checkerCommand?.DoCommand()
        End Sub

        Private Sub OnPopupMenuShowSpellingForm_ItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs)
            spellChecker.Check(Me.ActiveControl)
        End Sub

        Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Dim activeSheet = spreadsheetControl.ActiveWorksheet
            activeSheet.Columns(1).WidthInPixels = 150
            spreadsheetControl.SelectedCell = activeSheet("B2")
            spreadsheetControl.SelectedCell.Value = "Missspelled wods"
        End Sub
    End Class
End Namespace
