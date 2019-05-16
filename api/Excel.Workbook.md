---
title: Workbook object (Excel)
keywords: vbaxl10.chm198072
f1_keywords:
- vbaxl10.chm198072
ms.prod: excel
api_name:
- Excel.Workbook
ms.assetid: 8c00aa60-c974-eed3-0812-3c9625eb0d4c
ms.date: 05/15/2019
localization_priority: Normal
---


# Workbook object (Excel)

Represents a Microsoft Excel workbook.


## Remarks

The **Workbook** object is a member of the **[Workbooks](Excel.Workbooks.md)** collection. The **Workbooks** collection contains all the **Workbook** objects currently open in Microsoft Excel.

The **[ThisWorkbook](Excel.Application.ThisWorkbook.md)** property of the **Application** object returns the workbook where the Visual Basic code is running. In most cases, this is the same as the active workbook. However, if the Visual Basic code is part of an add-in, the **ThisWorkbook** property won't return the active workbook. In this case, the active workbook is the workbook calling the add-in, whereas the **ThisWorkbook** property returns the add-in workbook.

If you are creating an add-in from your Visual Basic code, you should use the **ThisWorkbook** property to qualify any statement that must be run on the workbook that you compile into the add-in.


## Example

Use **[Workbooks](Excel.Application.Workbooks.md)** (_index_), where _index_ is the workbook name or index number, to return a single **Workbook** object. The following example activates workbook one.

```vb
Workbooks(1).Activate
```

The index number denotes the order in which the workbooks were opened or created. `Workbooks(1)` is the first workbook created, and `Workbooks(Workbooks.Count)` is the last one created. Activating a workbook doesn't change its index number. All workbooks are included in the index count, even if they are hidden.

<br/>

The **[Name](Excel.Workbook.Name.md)** property returns the workbook name. You cannot set the name by using this property; if you need to change the name, use the **[SaveAs](Excel.Workbook.SaveAs.md)** method to save the workbook under a different name. 

The following example activates Sheet1 in the workbook named Cogs.xls (the workbook must already be open in Microsoft Excel).

```vb
Workbooks("Cogs.xls").Worksheets("Sheet1").Activate
```

<br/>

The **[ActiveWorkbook](Excel.Application.ActiveWorkbook.md)** property of the **Application** object returns the workbook that's currently active. The following example sets the name of the author for the active workbook.

```vb
ActiveWorkbook.Author = "Jean Selva"
```

<br/>

This example emails a worksheet tab from the active workbook by using a specified email address and subject. To run this code, the active worksheet must contain the email address in cell A1, the subject in cell B1, and the name of the worksheet to send in cell C1.

```vb
Sub SendTab()
   'Declare and initialize your variables, and turn off screen updating.
   Dim wks As Worksheet
   Application.ScreenUpdating = False
   Set wks = ActiveSheet
   
   'Copy the target worksheet, specified in cell C1, to the clipboard.
   Worksheets(Range("C1").Value).Copy
   
   'Send the content in the clipboard to the email account specified in cell A1,
   'using the subject line specified in cell B1.
   ActiveWorkbook.SendMail wks.Range("A1").Value, wks.Range("B1").Value
   
   'Do not save changes, and turn screen updating back on.
   ActiveWorkbook.Close savechanges:=False
   Application.ScreenUpdating = True
End Sub
```


## Events

- [Activate](Excel.Workbook.Activate(even).md)
- [AddinInstall](Excel.Workbook.AddinInstall.md)
- [AddinUninstall](Excel.Workbook.AddinUninstall.md)
- [AfterRemoteChange](Excel.Workbook.AfterRemoteChange.md)
- [AfterSave](Excel.Workbook.AfterSave.md)
- [AfterXmlExport](Excel.Workbook.AfterXmlExport.md)
- [AfterXmlImport](Excel.Workbook.AfterXmlImport.md)
- [BeforeClose](Excel.Workbook.BeforeClose.md)
- [BeforePrint](Excel.Workbook.BeforePrint.md)
- [BeforeRemoteChange](Excel.Workbook.BeforeRemoteChange.md)
- [BeforeSave](Excel.Workbook.BeforeSave.md)
- [BeforeXmlExport](Excel.Workbook.BeforeXmlExport.md)
- [BeforeXmlImport](Excel.Workbook.BeforeXmlImport.md)
- [Deactivate](Excel.Workbook.Deactivate.md)
- [ModelChange](Excel.workbook.modelchange.md)
- [NewChart](Excel.Workbook.NewChart.md)
- [NewSheet](Excel.Workbook.NewSheet.md)
- [Open](Excel.Workbook.Open.md)
- [PivotTableCloseConnection](Excel.workbook.pivottablecloseconnection.md)
- [PivotTableOpenConnection](Excel.workbook.pivottableopenconnection.md)
- [RowsetComplete](Excel.Workbook.RowsetComplete.md)
- [SheetActivate](Excel.Workbook.SheetActivate.md)
- [SheetBeforeDelete](Excel.workbook.sheetbeforedelete.md)
- [SheetBeforeDoubleClick](Excel.Workbook.SheetBeforeDoubleClick.md)
- [SheetBeforeRightClick](Excel.Workbook.SheetBeforeRightClick.md)
- [SheetCalculate](Excel.Workbook.SheetCalculate.md)
- [SheetChange](Excel.Workbook.SheetChange.md)
- [SheetDeactivate](Excel.Workbook.SheetDeactivate.md)
- [SheetFollowHyperlink](Excel.Workbook.SheetFollowHyperlink.md)
- [SheetLensGalleryRenderComplete](Excel.workbook.sheetlensgalleryrendercomplete.md)
- [SheetPivotTableAfterValueChange](Excel.Workbook.SheetPivotTableAfterValueChange.md)
- [SheetPivotTableBeforeAllocateChanges](Excel.Workbook.SheetPivotTableBeforeAllocateChanges.md)
- [SheetPivotTableBeforeCommitChanges](Excel.Workbook.SheetPivotTableBeforeCommitChanges.md)
- [SheetPivotTableBeforeDiscardChanges](Excel.Workbook.SheetPivotTableBeforeDiscardChanges.md)
- [SheetPivotTableChangeSync](Excel.Workbook.SheetPivotTableChangeSync.md)
- [SheetPivotTableUpdate](Excel.Workbook.SheetPivotTableUpdate.md)
- [SheetSelectionChange](Excel.Workbook.SheetSelectionChange.md)
- [SheetTableUpdate](Excel.workbook.sheettableupdate.md)
- [Sync](Excel.workbook.sync(event).md)
- [WindowActivate](Excel.Workbook.WindowActivate.md)
- [WindowDeactivate](Excel.Workbook.WindowDeactivate.md)
- [WindowResize](Excel.Workbook.WindowResize.md)

## Methods

- [AcceptAllChanges](Excel.Workbook.AcceptAllChanges.md)
- [Activate](Excel.Workbook.Activate(method).md)
- [AddToFavorites](Excel.Workbook.AddToFavorites.md)
- [ApplyTheme](Excel.Workbook.ApplyTheme.md)
- [BreakLink](Excel.Workbook.BreakLink.md)
- [CanCheckIn](Excel.Workbook.CanCheckIn.md)
- [ChangeFileAccess](Excel.Workbook.ChangeFileAccess.md)
- [ChangeLink](Excel.Workbook.ChangeLink.md)
- [CheckIn](Excel.Workbook.CheckIn.md)
- [CheckInWithVersion](Excel.Workbook.CheckInWithVersion.md)
- [Close](Excel.Workbook.Close.md)
- [ConvertComments](Excel.Workbook.ConvertComments.md)
- [CreateForecastSheet](Excel.workbook.createforecastsheet.md)
- [DeleteNumberFormat](Excel.Workbook.DeleteNumberFormat.md)
- [EnableConnections](Excel.Workbook.EnableConnections.md)
- [EndReview](Excel.Workbook.EndReview.md)
- [ExclusiveAccess](Excel.Workbook.ExclusiveAccess.md)
- [ExportAsFixedFormat](Excel.Workbook.ExportAsFixedFormat.md)
- [FollowHyperlink](Excel.Workbook.FollowHyperlink.md)
- [ForwardMailer](Excel.Workbook.ForwardMailer.md)
- [GetWorkflowTasks](Excel.Workbook.GetWorkflowTasks.md)
- [GetWorkflowTemplates](Excel.Workbook.GetWorkflowTemplates.md)
- [HighlightChangesOptions](Excel.Workbook.HighlightChangesOptions.md)
- [LinkInfo](Excel.Workbook.LinkInfo.md)
- [LinkSources](Excel.Workbook.LinkSources.md)
- [LockServerFile](Excel.Workbook.LockServerFile.md)
- [MergeWorkbook](Excel.Workbook.MergeWorkbook.md)
- [NewWindow](Excel.Workbook.NewWindow.md)
- [OpenLinks](Excel.Workbook.OpenLinks.md)
- [PivotCaches](Excel.Workbook.PivotCaches.md)
- [Post](Excel.Workbook.Post.md)
- [PrintOut](Excel.Workbook.PrintOut.md)
- [PrintPreview](Excel.Workbook.PrintPreview.md)
- [Protect](Excel.Workbook.Protect.md)
- [ProtectSharing](Excel.Workbook.ProtectSharing.md)
- [PublishToDocs](Excel.workbook.publishtodocs.md)
- [PurgeChangeHistoryNow](Excel.Workbook.PurgeChangeHistoryNow.md)
- [RefreshAll](Excel.Workbook.RefreshAll.md)
- [RejectAllChanges](Excel.Workbook.RejectAllChanges.md)
- [ReloadAs](Excel.Workbook.ReloadAs.md)
- [RemoveDocumentInformation](Excel.Workbook.RemoveDocumentInformation.md)
- [RemoveUser](Excel.Workbook.RemoveUser.md)
- [Reply](Excel.Workbook.Reply.md)
- [ReplyAll](Excel.Workbook.ReplyAll.md)
- [ReplyWithChanges](Excel.Workbook.ReplyWithChanges.md)
- [ResetColors](Excel.Workbook.ResetColors.md)
- [RunAutoMacros](Excel.Workbook.RunAutoMacros.md)
- [Save](Excel.Workbook.Save.md)
- [SaveAs](Excel.Workbook.SaveAs.md)
- [SaveAsXMLData](Excel.Workbook.SaveAsXMLData.md)
- [SaveCopyAs](Excel.Workbook.SaveCopyAs.md)
- [SendFaxOverInternet](Excel.Workbook.SendFaxOverInternet.md)
- [SendForReview](Excel.Workbook.SendForReview.md)
- [SendMail](Excel.Workbook.SendMail.md)
- [SendMailer](Excel.Workbook.SendMailer.md)
- [SetLinkOnData](Excel.Workbook.SetLinkOnData.md)
- [SetPasswordEncryptionOptions](Excel.Workbook.SetPasswordEncryptionOptions.md)
- [ToggleFormsDesign](Excel.Workbook.ToggleFormsDesign.md)
- [Unprotect](Excel.Workbook.Unprotect.md)
- [UnprotectSharing](Excel.Workbook.UnprotectSharing.md)
- [UpdateFromFile](Excel.Workbook.UpdateFromFile.md)
- [UpdateLink](Excel.Workbook.UpdateLink.md)
- [WebPagePreview](Excel.Workbook.WebPagePreview.md)
- [XmlImport](Excel.Workbook.XmlImport.md)
- [XmlImportXml](Excel.Workbook.XmlImportXml.md)

## Properties

- [AccuracyVersion](Excel.Workbook.AccuracyVersion.md)
- [ActiveChart](Excel.Workbook.ActiveChart.md)
- [ActiveSheet](Excel.Workbook.ActiveSheet.md)
- [ActiveSlicer](Excel.Workbook.ActiveSlicer.md)
- [Application](Excel.Workbook.Application.md)
- [AutoSaveOn](Excel.Workbook.AutoSaveOn.md)
- [AutoUpdateFrequency](Excel.Workbook.AutoUpdateFrequency.md)
- [AutoUpdateSaveChanges](Excel.Workbook.AutoUpdateSaveChanges.md)
- [BuiltinDocumentProperties](Excel.Workbook.BuiltinDocumentProperties.md)
- [CalculationVersion](Excel.Workbook.CalculationVersion.md)
- [CaseSensitive](Excel.workbook.casesensitive.md)
- [ChangeHistoryDuration](Excel.Workbook.ChangeHistoryDuration.md)
- [ChartDataPointTrack](Excel.workbook.chartdatapointtrack.md)
- [Charts](Excel.Workbook.Charts.md)
- [CheckCompatibility](Excel.Workbook.CheckCompatibility.md)
- [CodeName](Excel.Workbook.CodeName.md)
- [Colors](Excel.Workbook.Colors.md)
- [CommandBars](Excel.Workbook.CommandBars.md)
- [ConflictResolution](Excel.Workbook.ConflictResolution.md)
- [Connections](Excel.Workbook.Connections.md)
- [ConnectionsDisabled](Excel.Workbook.ConnectionsDisabled.md)
- [Container](Excel.Workbook.Container.md)
- [ContentTypeProperties](Excel.Workbook.ContentTypeProperties.md)
- [CreateBackup](Excel.Workbook.CreateBackup.md)
- [Creator](Excel.Workbook.Creator.md)
- [CustomDocumentProperties](Excel.Workbook.CustomDocumentProperties.md)
- [CustomViews](Excel.Workbook.CustomViews.md)
- [CustomXMLParts](Excel.Workbook.CustomXMLParts.md)
- [Date1904](Excel.Workbook.Date1904.md)
- [DefaultPivotTableStyle](Excel.Workbook.DefaultPivotTableStyle.md)
- [DefaultSlicerStyle](Excel.Workbook.DefaultSlicerStyle.md)
- [DefaultTableStyle](Excel.Workbook.DefaultTableStyle.md)
- [DefaultTimelineStyle](Excel.workbook.defaulttimelinestyle.md)
- [DisplayDrawingObjects](Excel.Workbook.DisplayDrawingObjects.md)
- [DisplayInkComments](Excel.Workbook.DisplayInkComments.md)
- [DocumentInspectors](Excel.Workbook.DocumentInspectors.md)
- [DocumentLibraryVersions](Excel.Workbook.DocumentLibraryVersions.md)
- [DoNotPromptForConvert](Excel.Workbook.DoNotPromptForConvert.md)
- [EnableAutoRecover](Excel.Workbook.EnableAutoRecover.md)
- [EncryptionProvider](Excel.Workbook.EncryptionProvider.md)
- [EnvelopeVisible](Excel.Workbook.EnvelopeVisible.md)
- [Excel4IntlMacroSheets](Excel.Workbook.Excel4IntlMacroSheets.md)
- [Excel4MacroSheets](Excel.Workbook.Excel4MacroSheets.md)
- [Excel8CompatibilityMode](Excel.Workbook.Excel8CompatibilityMode.md)
- [FileFormat](Excel.Workbook.FileFormat.md)
- [Final](Excel.Workbook.Final.md)
- [ForceFullCalculation](Excel.Workbook.ForceFullCalculation.md)
- [FullName](Excel.Workbook.FullName.md)
- [FullNameURLEncoded](Excel.Workbook.FullNameURLEncoded.md)
- [HasPassword](Excel.Workbook.HasPassword.md)
- [HasVBProject](Excel.Workbook.HasVBProject.md)
- [HighlightChangesOnScreen](Excel.Workbook.HighlightChangesOnScreen.md)
- [IconSets](Excel.Workbook.IconSets.md)
- [InactiveListBorderVisible](Excel.Workbook.InactiveListBorderVisible.md)
- [IsAddin](Excel.Workbook.IsAddin.md)
- [IsInplace](Excel.Workbook.IsInplace.md)
- [KeepChangeHistory](Excel.Workbook.KeepChangeHistory.md)
- [ListChangesOnNewSheet](Excel.Workbook.ListChangesOnNewSheet.md)
- [Mailer](Excel.Workbook.Mailer.md)
- [Model](Excel.workbook.model.md)
- [MultiUserEditing](Excel.Workbook.MultiUserEditing.md)
- [Name](Excel.Workbook.Name.md)
- [Names](Excel.Workbook.Names.md)
- [Parent](Excel.Workbook.Parent.md)
- [Password](Excel.Workbook.Password.md)
- [PasswordEncryptionAlgorithm](Excel.Workbook.PasswordEncryptionAlgorithm.md)
- [PasswordEncryptionFileProperties](Excel.Workbook.PasswordEncryptionFileProperties.md)
- [PasswordEncryptionKeyLength](Excel.Workbook.PasswordEncryptionKeyLength.md)
- [PasswordEncryptionProvider](Excel.Workbook.PasswordEncryptionProvider.md)
- [Path](Excel.Workbook.Path.md)
- [Permission](Excel.Workbook.Permission.md)
- [PersonalViewListSettings](Excel.Workbook.PersonalViewListSettings.md)
- [PersonalViewPrintSettings](Excel.Workbook.PersonalViewPrintSettings.md)
- [PivotTables](Excel.workbook.pivottables.md)
- [PrecisionAsDisplayed](Excel.Workbook.PrecisionAsDisplayed.md)
- [ProtectStructure](Excel.Workbook.ProtectStructure.md)
- [ProtectWindows](Excel.Workbook.ProtectWindows.md)
- [PublishObjects](Excel.Workbook.PublishObjects.md)
- [Queries](Excel.workbook.queries.md)
- [ReadOnly](Excel.Workbook.ReadOnly.md)
- [ReadOnlyRecommended](Excel.Workbook.ReadOnlyRecommended.md)
- [RemovePersonalInformation](Excel.Workbook.RemovePersonalInformation.md)
- [Research](Excel.Workbook.Research.md)
- [RevisionNumber](Excel.Workbook.RevisionNumber.md)
- [Saved](Excel.Workbook.Saved.md)
- [SaveLinkValues](Excel.Workbook.SaveLinkValues.md)
- [ServerPolicy](Excel.Workbook.ServerPolicy.md)
- [ServerViewableItems](Excel.Workbook.ServerViewableItems.md)
- [SharedWorkspace](Excel.Workbook.SharedWorkspace.md)
- [Sheets](Excel.Workbook.Sheets.md)
- [ShowConflictHistory](Excel.Workbook.ShowConflictHistory.md)
- [ShowPivotChartActiveFields](Excel.Workbook.ShowPivotChartActiveFields.md)
- [ShowPivotTableFieldList](Excel.Workbook.ShowPivotTableFieldList.md)
- [Signatures](Excel.Workbook.Signatures.md)
- [SlicerCaches](Excel.Workbook.SlicerCaches.md)
- [SmartDocument](Excel.Workbook.SmartDocument.md)
- [Styles](Excel.Workbook.Styles.md)
- [Sync](Excel.Workbook.Sync.md)
- [TableStyles](Excel.Workbook.TableStyles.md)
- [TemplateRemoveExtData](Excel.Workbook.TemplateRemoveExtData.md)
- [Theme](Excel.Workbook.Theme.md)
- [UpdateLinks](Excel.Workbook.UpdateLinks.md)
- [UpdateRemoteReferences](Excel.Workbook.UpdateRemoteReferences.md)
- [UserStatus](Excel.Workbook.UserStatus.md)
- [UseWholeCellCriteria](Excel.workbook.usewholecellcriteria.md)
- [UseWildcards](Excel.workbook.usewildcards.md)
- [VBASigned](Excel.Workbook.VBASigned.md)
- [VBProject](Excel.Workbook.VBProject.md)
- [WebOptions](Excel.Workbook.WebOptions.md)
- [Windows](Excel.Workbook.Windows.md)
- [Worksheets](Excel.Workbook.Worksheets.md)
- [WritePassword](Excel.Workbook.WritePassword.md)
- [WriteReserved](Excel.Workbook.WriteReserved.md)
- [WriteReservedBy](Excel.Workbook.WriteReservedBy.md)
- [XmlMaps](Excel.Workbook.XmlMaps.md)
- [XmlNamespaces](Excel.Workbook.XmlNamespaces.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
