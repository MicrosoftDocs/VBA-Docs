---
title: Application object (Excel)
keywords: vbaxl10.chm182073
f1_keywords:
- vbaxl10.chm182073
ms.prod: excel
api_name:
- Excel.Application
ms.assetid: 19b73597-5cf9-4f56-8227-b5211f657f6f
ms.date: 06/08/2017
localization_priority: Priority
---

# Application object (Excel)

Represents the entire Microsoft Excel application.


## Example

Use the **Application** property to return the **Application** object. The following example applies the **Windows** property to the **Application** object.

```vb
Application.Windows("book1.xls").Activate
```

<br/>

The following example creates an Excel workbook object in another application and then opens a workbook in Excel.

```vb
Set xl = CreateObject("Excel.Sheet") 
xl.Application.Workbooks.Open "newbook.xls"
```

<br/>

Many of the properties and methods that return the most common user-interface objects, such as the active cell (**ActiveCell** property), can be used without the **Application** object qualifier. For example, instead of writing:

```vb
Application.ActiveCell.Font.Bold = True
```

You can write: 

```vb
ActiveCell.Font.Bold = True
```


## Remarks

The  **Application** object contains:

- Application-wide settings and options.
    
- Methods that return top-level objects, such as **[ActiveCell](Excel.Application.ActiveCell.md)**, **[ActiveSheet](Excel.Application.ActiveSheet.md)**, and so on.
    
## Events

|Name|
|:-----|
|[AfterCalculate](Excel.Application.AfterCalculate.md)|
|[NewWorkbook](Excel.Application.NewWorkbook(even).md)|
|[ProtectedViewWindowActivate](Excel.Application.ProtectedViewWindowActivate.md)|
|[ProtectedViewWindowBeforeClose](Excel.Application.ProtectedViewWindowBeforeClose.md)|
|[ProtectedViewWindowBeforeEdit](Excel.Application.ProtectedViewWindowBeforeEdit.md)|
|[ProtectedViewWindowDeactivate](Excel.Application.ProtectedViewWindowDeactivate.md)|
|[ProtectedViewWindowOpen](Excel.Application.ProtectedViewWindowOpen.md)|
|[ProtectedViewWindowResize](Excel.Application.ProtectedViewWindowResize.md)|
|[SheetActivate](Excel.Application.SheetActivate.md)|
|[SheetBeforeDelete](Excel.application.sheetbeforedelete.md)|
|[SheetBeforeDoubleClick](Excel.Application.SheetBeforeDoubleClick.md)|
|[SheetBeforeRightClick](Excel.Application.SheetBeforeRightClick.md)|
|[SheetCalculate](Excel.Application.SheetCalculate.md)|
|[SheetChange](Excel.Application.SheetChange.md)|
|[SheetDeactivate](Excel.Application.SheetDeactivate.md)|
|[SheetFollowHyperlink](Excel.Application.SheetFollowHyperlink.md)|
|[SheetLensGalleryRenderComplete](Excel.application.sheetlensgalleryrendercomplete.md)|
|[SheetPivotTableAfterValueChange](Excel.Application.SheetPivotTableAfterValueChange.md)|
|[SheetPivotTableBeforeAllocateChanges](Excel.Application.SheetPivotTableBeforeAllocateChanges.md)|
|[SheetPivotTableBeforeCommitChanges](Excel.Application.SheetPivotTableBeforeCommitChanges.md)|
|[SheetPivotTableBeforeDiscardChanges](Excel.Application.SheetPivotTableBeforeDiscardChanges.md)|
|[SheetPivotTableUpdate](Excel.Application.SheetPivotTableUpdate.md)|
|[SheetSelectionChange](Excel.Application.SheetSelectionChange.md)|
|[SheetTableUpdate](Excel.application.sheettableupdate.md)|
|[WindowActivate](Excel.Application.WindowActivate.md)|
|[WindowDeactivate](Excel.Application.WindowDeactivate.md)|
|[WindowResize](Excel.Application.WindowResize.md)|
|[WorkbookActivate](Excel.Application.WorkbookActivate.md)|
|[WorkbookAddinInstall](Excel.Application.WorkbookAddinInstall.md)|
|[WorkbookAddinUninstall](Excel.Application.WorkbookAddinUninstall.md)|
|[WorkbookAfterSave](overview/Excel.md)|
|[WorkbookAfterXmlExport](Excel.Application.WorkbookAfterXmlExport.md)|
|[WorkbookAfterXmlImport](Excel.Application.WorkbookAfterXmlImport.md)|
|[WorkbookBeforeClose](Excel.Application.WorkbookBeforeClose.md)|
|[WorkbookBeforePrint](Excel.Application.WorkbookBeforePrint.md)|
|[WorkbookBeforeSave](Excel.Application.WorkbookBeforeSave.md)|
|[WorkbookBeforeXmlExport](Excel.Application.WorkbookBeforeXmlExport.md)|
|[WorkbookBeforeXmlImport](Excel.Application.WorkbookBeforeXmlImport.md)|
|[WorkbookDeactivate](Excel.Application.WorkbookDeactivate.md)|
|[WorkbookModelChange](Excel.application.workbookmodelchange.md)|
|[WorkbookNewChart](Excel.Application.WorkbookNewChart.md)|
|[WorkbookNewSheet](Excel.Application.WorkbookNewSheet.md)|
|[WorkbookOpen](Excel.Application.WorkbookOpen.md)|
|[WorkbookPivotTableCloseConnection](Excel.Application.WorkbookPivotTableCloseConnection.md)|
|[WorkbookPivotTableOpenConnection](Excel.Application.WorkbookPivotTableOpenConnection.md)|
|[WorkbookRowsetComplete](Excel.Application.WorkbookRowsetComplete.md)|
|[WorkbookSync](Excel.Application.WorkbookSync.md)|

<br/>

## Methods

|Name|
|:-----|
|[ActivateMicrosoftApp](Excel.Application.ActivateMicrosoftApp.md)|
|[AddCustomList](Excel.Application.AddCustomList.md)|
|[Calculate](Excel.Application.Calculate.md)|
|[CalculateFull](Excel.Application.CalculateFull.md)|
|[CalculateFullRebuild](Excel.Application.CalculateFullRebuild.md)|
|[CalculateUntilAsyncQueriesDone](Excel.Application.CalculateUntilAsyncQueriesDone.md)|
|[CentimetersToPoints](Excel.Application.CentimetersToPoints.md)|
|[CheckAbort](Excel.Application.CheckAbort.md)|
|[CheckSpelling](Excel.Application.CheckSpelling.md)|
|[ConvertFormula](Excel.Application.ConvertFormula.md)|
|[DDEExecute](Excel.Application.DDEExecute.md)|
|[DDEInitiate](Excel.Application.DDEInitiate.md)|
|[DDEPoke](Excel.Application.DDEPoke.md)|
|[DDERequest](Excel.Application.DDERequest.md)|
|[DDETerminate](Excel.Application.DDETerminate.md)|
|[DeleteCustomList](Excel.Application.DeleteCustomList.md)|
|[DisplayXMLSourcePane](Excel.Application.DisplayXMLSourcePane.md)|
|[DoubleClick](Excel.Application.DoubleClick.md)|
|[Evaluate](Excel.Application.Evaluate.md)|
|[ExecuteExcel4Macro](Excel.Application.ExecuteExcel4Macro.md)|
|[FindFile](Excel.Application.FindFile.md)|
|[GetCustomListContents](Excel.Application.GetCustomListContents.md)|
|[GetCustomListNum](Excel.Application.GetCustomListNum.md)|
|[GetOpenFilename](Excel.Application.GetOpenFilename.md)|
|[GetPhonetic](Excel.Application.GetPhonetic.md)|
|[GetSaveAsFilename](Excel.Application.GetSaveAsFilename.md)|
|[Goto](Excel.Application.Goto.md)|
|[Help](Excel.Application.Help.md)|
|[InchesToPoints](Excel.Application.InchesToPoints.md)|
|[InputBox](Excel.Application.InputBox.md)|
|[Intersect](Excel.Application.Intersect.md)|
|[MacroOptions](Excel.Application.MacroOptions.md)|
|[MailLogoff](Excel.Application.MailLogoff.md)|
|[MailLogon](Excel.Application.MailLogon.md)|
|[NextLetter](Excel.Application.NextLetter.md)|
|[OnKey](Excel.Application.OnKey.md)|
|[OnRepeat](Excel.Application.OnRepeat.md)|
|[OnTime](Excel.Application.OnTime.md)|
|[OnUndo](Excel.Application.OnUndo.md)|
|[Quit](Excel.Application.Quit.md)|
|[RecordMacro](Excel.Application.RecordMacro.md)|
|[RegisterXLL](Excel.Application.RegisterXLL.md)|
|[Repeat](Excel.Application.Repeat.md)|
|[Run](Excel.Application.Run.md)|
|[SendKeys](Excel.Application.SendKeys.md)|
|[SharePointVersion](Excel.Application.SharePointVersion.md)|
|[Undo](Excel.Application.Undo.md)|
|[Union](Excel.Application.Union.md)|
|[Volatile](Excel.Application.Volatile.md)|
|[Wait](Excel.Application.Wait.md)|

<br/>

## Properties

|Name|
|:-----|
|[ActiveCell](Excel.Application.ActiveCell.md)|
|[ActiveChart](Excel.Application.ActiveChart.md)|
|[ActiveEncryptionSession](Excel.Application.ActiveEncryptionSession.md)|
|[ActivePrinter](Excel.Application.ActivePrinter.md)|
|[ActiveProtectedViewWindow](Excel.Application.ActiveProtectedViewWindow.md)|
|[ActiveSheet](Excel.Application.ActiveSheet.md)|
|[ActiveWindow](Excel.Application.ActiveWindow.md)|
|[ActiveWorkbook](Excel.Application.ActiveWorkbook.md)|
|[AddIns](Excel.Application.AddIns.md)|
|[AddIns2](Excel.Application.AddIns2.md)|
|[AlertBeforeOverwriting](Excel.Application.AlertBeforeOverwriting.md)|
|[AltStartupPath](Excel.Application.AltStartupPath.md)|
|[AlwaysUseClearType](Excel.Application.AlwaysUseClearType.md)|
|[Application](Excel.Application.Application.md)|
|[ArbitraryXMLSupportAvailable](Excel.Application.ArbitraryXMLSupportAvailable.md)|
|[AskToUpdateLinks](Excel.Application.AskToUpdateLinks.md)|
|[Assistance](Excel.Application.Assistance.md)|
|[AutoCorrect](Excel.Application.AutoCorrect.md)|
|[AutoFormatAsYouTypeReplaceHyperlinks](Excel.Application.AutoFormatAsYouTypeReplaceHyperlinks.md)|
|[AutomationSecurity](Excel.Application.AutomationSecurity.md)|
|[AutoPercentEntry](Excel.Application.AutoPercentEntry.md)|
|[AutoRecover](Excel.Application.AutoRecover.md)|
|[Build](Excel.Application.Build.md)|
|[CalculateBeforeSave](Excel.Application.CalculateBeforeSave.md)|
|[Calculation](Excel.Application.Calculation.md)|
|[CalculationInterruptKey](Excel.Application.CalculationInterruptKey.md)|
|[CalculationState](Excel.Application.CalculationState.md)|
|[CalculationVersion](Excel.Application.CalculationVersion.md)|
|[Caller](Excel.Application.Caller.md)|
|[CanPlaySounds](Excel.Application.CanPlaySounds.md)|
|[CanRecordSounds](Excel.Application.CanRecordSounds.md)|
|[Caption](Excel.Application.Caption.md)|
|[CellDragAndDrop](Excel.Application.CellDragAndDrop.md)|
|[Cells](Excel.Application.Cells.md)|
|[ChartDataPointTrack](Excel.application.chartdatapointtrack.md)|
|[Charts](Excel.Application.Charts.md)|
|[ClipboardFormats](Excel.Application.ClipboardFormats.md)|
|[ClusterConnector](Excel.Application.ClusterConnector.md)|
|[Columns](Excel.Application.Columns.md)|
|[COMAddIns](Excel.Application.COMAddIns.md)|
|[CommandBars](Excel.Application.CommandBars.md)|
|[CommandUnderlines](Excel.Application.CommandUnderlines.md)|
|[ConstrainNumeric](Excel.Application.ConstrainNumeric.md)|
|[ControlCharacters](Excel.Application.ControlCharacters.md)|
|[CopyObjectsWithCells](Excel.Application.CopyObjectsWithCells.md)|
|[Creator](Excel.Application.Creator.md)|
|[Cursor](Excel.Application.Cursor.md)|
|[CursorMovement](Excel.Application.CursorMovement.md)|
|[CustomListCount](Excel.Application.CustomListCount.md)|
|[CutCopyMode](Excel.Application.CutCopyMode.md)|
|[DataEntryMode](Excel.Application.DataEntryMode.md)|
|[DDEAppReturnCode](Excel.Application.DDEAppReturnCode.md)|
|[DecimalSeparator](Excel.Application.DecimalSeparator.md)|
|[DefaultFilePath](Excel.Application.DefaultFilePath.md)|
|[DefaultSaveFormat](Excel.Application.DefaultSaveFormat.md)|
|[DefaultSheetDirection](Excel.Application.DefaultSheetDirection.md)|
|[DefaultWebOptions](Excel.Application.DefaultWebOptions.md)|
|[DeferAsyncQueries](Excel.Application.DeferAsyncQueries.md)|
|[Dialogs](Excel.Application.Dialogs.md)|
|[DisplayAlerts](Excel.Application.DisplayAlerts.md)|
|[DisplayClipboardWindow](Excel.Application.DisplayClipboardWindow.md)|
|[DisplayCommentIndicator](Excel.Application.DisplayCommentIndicator.md)|
|[DisplayDocumentActionTaskPane](Excel.Application.DisplayDocumentActionTaskPane.md)|
|[DisplayDocumentInformationPanel](Excel.Application.DisplayDocumentInformationPanel.md)|
|[DisplayExcel4Menus](Excel.Application.DisplayExcel4Menus.md)|
|[DisplayFormulaAutoComplete](Excel.Application.DisplayFormulaAutoComplete.md)|
|[DisplayFormulaBar](Excel.Application.DisplayFormulaBar.md)|
|[DisplayFullScreen](Excel.Application.DisplayFullScreen.md)|
|[DisplayFunctionToolTips](Excel.Application.DisplayFunctionToolTips.md)|
|[DisplayInsertOptions](Excel.Application.DisplayInsertOptions.md)|
|[DisplayNoteIndicator](Excel.Application.DisplayNoteIndicator.md)|
|[DisplayPasteOptions](Excel.Application.DisplayPasteOptions.md)|
|[DisplayRecentFiles](Excel.Application.DisplayRecentFiles.md)|
|[DisplayScrollBars](Excel.Application.DisplayScrollBars.md)|
|[DisplayStatusBar](Excel.Application.DisplayStatusBar.md)|
|[EditDirectlyInCell](Excel.Application.EditDirectlyInCell.md)|
|[EnableAnimations](Excel.Application.EnableAnimations.md)|
|[EnableAutoComplete](Excel.Application.EnableAutoComplete.md)|
|[EnableCancelKey](Excel.Application.EnableCancelKey.md)|
|[EnableCheckFileExtensions](Excel.application.enablecheckfileextensions.md)|
|[EnableEvents](Excel.Application.EnableEvents.md)|
|[EnableLargeOperationAlert](Excel.Application.EnableLargeOperationAlert.md)|
|[EnableLivePreview](Excel.Application.EnableLivePreview.md)|
|[EnableMacroAnimations](Excel.application.enablemacroanimations.md)|
|[EnableSound](Excel.Application.EnableSound.md)|
|[ErrorCheckingOptions](Excel.Application.ErrorCheckingOptions.md)|
|[Excel4IntlMacroSheets](Excel.Application.Excel4IntlMacroSheets.md)|
|[Excel4MacroSheets](Excel.Application.Excel4MacroSheets.md)|
|[ExtendList](Excel.Application.ExtendList.md)|
|[FeatureInstall](Excel.Application.FeatureInstall.md)|
|[FileConverters](Excel.Application.FileConverters.md)|
|[FileDialog](Excel.Application.FileDialog.md)|
|[FileExportConverters](Excel.Application.FileExportConverters.md)|
|[FileValidation](Excel.Application.FileValidation.md)|
|[FileValidationPivot](Excel.Application.FileValidationPivot.md)|
|[FindFormat](Excel.Application.FindFormat.md)|
|[FixedDecimal](Excel.Application.FixedDecimal.md)|
|[FixedDecimalPlaces](Excel.Application.FixedDecimalPlaces.md)|
|[FlashFill](Excel.application.flashfill.md)|
|[FlashFillMode](Excel.application.flashfillmode.md)|
|[FormulaBarHeight](Excel.Application.FormulaBarHeight.md)|
|[GenerateGetPivotData](Excel.Application.GenerateGetPivotData.md)|
|[GenerateTableRefs](Excel.Application.GenerateTableRefs.md)|
|[Height](Excel.Application.Height.md)|
|[HighQualityModeForGraphics](Excel.Application.HighQualityModeForGraphics.md)|
|[Hinstance](Excel.Application.Hinstance.md)|
|[HinstancePtr](Excel.Application.HinstancePtr.md)|
|[Hwnd](Excel.Application.Hwnd.md)|
|[IgnoreRemoteRequests](Excel.Application.IgnoreRemoteRequests.md)|
|[Interactive](Excel.Application.Interactive.md)|
|[International](Excel.Application.International.md)|
|[IsSandboxed](Excel.Application.IsSandboxed.md)|
|[Iteration](Excel.Application.Iteration.md)|
|[LanguageSettings](Excel.Application.LanguageSettings.md)|
|[LargeOperationCellThousandCount](Excel.Application.LargeOperationCellThousandCount.md)|
|[Left](Excel.Application.Left.md)|
|[LibraryPath](Excel.Application.LibraryPath.md)|
|[MailSession](Excel.Application.MailSession.md)|
|[MailSystem](Excel.Application.MailSystem.md)|
|[MapPaperSize](Excel.Application.MapPaperSize.md)|
|[MathCoprocessorAvailable](Excel.Application.MathCoprocessorAvailable.md)|
|[MaxChange](Excel.Application.MaxChange.md)|
|[MaxIterations](Excel.Application.MaxIterations.md)|
|[MeasurementUnit](Excel.Application.MeasurementUnit.md)|
|[MergeInstances](Excel.application.mergeinstances.md)|
|[MouseAvailable](Excel.Application.MouseAvailable.md)|
|[MoveAfterReturn](Excel.Application.MoveAfterReturn.md)|
|[MoveAfterReturnDirection](Excel.Application.MoveAfterReturnDirection.md)|
|[MultiThreadedCalculation](Excel.Application.MultiThreadedCalculation.md)|
|[Name](Excel.Application.Name.md)|
|[Names](Excel.Application.Names.md)|
|[NetworkTemplatesPath](Excel.Application.NetworkTemplatesPath.md)|
|[NewWorkbook](Excel.Application.NewWorkbook(property).md)|
|[ODBCErrors](Excel.Application.ODBCErrors.md)|
|[ODBCTimeout](Excel.Application.ODBCTimeout.md)|
|[OLEDBErrors](Excel.Application.OLEDBErrors.md)|
|[OnWindow](Excel.Application.OnWindow.md)|
|[OperatingSystem](Excel.Application.OperatingSystem.md)|
|[OrganizationName](Excel.Application.OrganizationName.md)|
|[Parent](Excel.Application.Parent.md)|
|[Path](Excel.Application.Path.md)|
|[PathSeparator](Excel.Application.PathSeparator.md)|
|[PivotTableSelection](Excel.Application.PivotTableSelection.md)|
|[PreviousSelections](Excel.Application.PreviousSelections.md)|
|[PrintCommunication](Excel.Application.PrintCommunication.md)|
|[ProductCode](Excel.Application.ProductCode.md)|
|[PromptForSummaryInfo](Excel.Application.PromptForSummaryInfo.md)|
|[ProtectedViewWindows](Excel.Application.ProtectedViewWindows.md)|
|[QuickAnalysis](Excel.application.quickanalysis.md)|
|[Range](Excel.Application.Range.md)|
|[Ready](Excel.Application.Ready.md)|
|[RecentFiles](Excel.Application.RecentFiles.md)|
|[RecordRelative](Excel.Application.RecordRelative.md)|
|[ReferenceStyle](Excel.Application.ReferenceStyle.md)|
|[RegisteredFunctions](Excel.Application.RegisteredFunctions.md)|
|[ReplaceFormat](Excel.Application.ReplaceFormat.md)|
|[RollZoom](Excel.Application.RollZoom.md)|
|[Rows](Excel.Application.Rows.md)|
|[RTD](Excel.Application.RTD.md)|
|[ScreenUpdating](Excel.Application.ScreenUpdating.md)|
|[Selection](Excel.Application.Selection.md)|
|[Sheets](Excel.Application.Sheets.md)|
|[SheetsInNewWorkbook](Excel.Application.SheetsInNewWorkbook.md)|
|[ShowChartTipNames](Excel.Application.ShowChartTipNames.md)|
|[ShowChartTipValues](Excel.Application.ShowChartTipValues.md)|
|[ShowDevTools](Excel.Application.ShowDevTools.md)|
|[ShowMenuFloaties](Excel.Application.ShowMenuFloaties.md)|
|[ShowQuickAnalysis](Excel.application.showquickanalysis.md)|
|[ShowSelectionFloaties](Excel.Application.ShowSelectionFloaties.md)|
|[ShowStartupDialog](Excel.Application.ShowStartupDialog.md)|
|[ShowToolTips](Excel.Application.ShowToolTips.md)|
|[SmartArtColors](Excel.Application.SmartArtColors.md)|
|[SmartArtLayouts](Excel.Application.SmartArtLayouts.md)|
|[SmartArtQuickStyles](Excel.Application.SmartArtQuickStyles.md)|
|[Speech](Excel.Application.Speech.md)|
|[SpellingOptions](Excel.Application.SpellingOptions.md)|
|[StandardFont](Excel.Application.StandardFont.md)|
|[StandardFontSize](Excel.Application.StandardFontSize.md)|
|[StartupPath](Excel.Application.StartupPath.md)|
|[StatusBar](Excel.Application.StatusBar.md)|
|[TemplatesPath](Excel.Application.TemplatesPath.md)|
|[ThisCell](Excel.Application.ThisCell.md)|
|[ThisWorkbook](Excel.Application.ThisWorkbook.md)|
|[ThousandsSeparator](Excel.Application.ThousandsSeparator.md)|
|[Top](Excel.Application.Top.md)|
|[TransitionMenuKey](Excel.Application.TransitionMenuKey.md)|
|[TransitionMenuKeyAction](Excel.Application.TransitionMenuKeyAction.md)|
|[TransitionNavigKeys](Excel.Application.TransitionNavigKeys.md)|
|[UsableHeight](Excel.Application.UsableHeight.md)|
|[UsableWidth](Excel.Application.UsableWidth.md)|
|[UseClusterConnector](Excel.Application.UseClusterConnector.md)|
|[UsedObjects](Excel.Application.UsedObjects.md)|
|[UserControl](Excel.Application.UserControl.md)|
|[UserLibraryPath](Excel.Application.UserLibraryPath.md)|
|[UserName](Excel.Application.UserName.md)|
|[UseSystemSeparators](Excel.Application.UseSystemSeparators.md)|
|[Value](Excel.Application.Value.md)|
|[VBE](Excel.Application.VBE.md)|
|[Version](Excel.Application.Version.md)|
|[Visible](Excel.Application.Visible.md)|
|[WarnOnFunctionNameConflict](Excel.Application.WarnOnFunctionNameConflict.md)|
|[Watches](Excel.Application.Watches.md)|
|[Width](Excel.Application.Width.md)|
|[Windows](Excel.Application.Windows.md)|
|[WindowsForPens](Excel.Application.WindowsForPens.md)|
|[WindowState](Excel.Application.WindowState.md)|
|[Workbooks](Excel.Application.Workbooks.md)|
|[WorksheetFunction](Excel.Application.WorksheetFunction.md)|
|[Worksheets](Excel.Application.Worksheets.md)|


<br/>

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]