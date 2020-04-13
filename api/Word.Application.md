---
title: Application object (Word)
keywords: vbawd10.chm2416
f1_keywords:
- vbawd10.chm2416
ms.prod: word
api_name:
- Word.Application
ms.assetid: d1cf6f8f-4e88-bf01-93b4-90a83f79cb44
ms.date: 06/08/2017
localization_priority: Normal
---


# Application object (Word)

Represents the Microsoft Word application. The **Application** object includes properties and methods that return top-level objects. For example, the **[ActiveDocument](./Word.Application.ActiveDocument.md)** property returns a **[Document](Word.Document.md)** object.


## Remarks

Use the **Application** property to return the **Application** object. The following example displays the user name for Word.


```vb
MsgBox Application.UserName
```

Many of the properties and methods that return the most common user-interface objects—such as the active document (**ActiveDocument** property)—can be used without the **Application** object qualifier. For example, instead of writing `Application.ActiveDocument.PrintOut`, you can write  `ActiveDocument.PrintOut`. Properties and methods that can be used without the **Application** object qualifier are considered "global." To view the global properties and methods in the **Object Browser**, click `<globals>` at the top of the list in the **Classes** box. (Also see the **[Global](./Word.Global.md)** object.)

Remarks

To use Automation (formerly OLE Automation) to control Word from another application, use the Microsoft Visual Basic  **CreateObject** or **GetObject** function to return a Word **Application** object. The following Microsoft Excel example starts Word (if it is not already running) and opens an existing document.




```vb
Set wrd = GetObject(, "Word.Application") 
wrd.Visible = True 
wrd.Documents.Open "C:\My Documents\Temp.doc" 
Set wrd = Nothing
```


## Events

- [DocumentBeforeClose](Word.Application.DocumentBeforeClose.md)
- [DocumentBeforePrint](Word.Application.DocumentBeforePrint.md)
- [DocumentBeforeSave](Word.Application.DocumentBeforeSave.md)
- [DocumentChange](Word.Application.DocumentChange.md)
- [DocumentOpen](Word.Application.DocumentOpen.md)
- [DocumentSync](Word.Application.DocumentSync.md)
- [EPostageInsert](Word.Application.EPostageInsert.md)
- [EPostageInsertEx](Word.Application.EPostageInsertEx.md)
- [EPostagePropertyDialog](Word.Application.EPostagePropertyDialog.md)
- [MailMergeAfterMerge](Word.Application.MailMergeAfterMerge.md)
- [MailMergeAfterRecordMerge](Word.Application.MailMergeAfterRecordMerge.md)
- [MailMergeBeforeMerge](Word.Application.MailMergeBeforeMerge.md)
- [MailMergeBeforeRecordMerge](Word.Application.MailMergeBeforeRecordMerge.md)
- [MailMergeDataSourceLoad](Word.Application.MailMergeDataSourceLoad.md)
- [MailMergeDataSourceValidate](Word.Application.MailMergeDataSourceValidate.md)
- [MailMergeDataSourceValidate2](Word.Application.MailMergeDataSourceValidate2.md)
- [MailMergeWizardSendToCustom](Word.Application.MailMergeWizardSendToCustom.md)
- [MailMergeWizardStateChange](Word.Application.MailMergeWizardStateChange.md)
- [NewDocument](Word.Application.NewDocument(even).md)
- [ProtectedViewWindowActivate](Word.Application.ProtectedViewWindowActivate.md)
- [ProtectedViewWindowBeforeClose](Word.Application.ProtectedViewWindowBeforeClose.md)
- [ProtectedViewWindowBeforeEdit](Word.Application.ProtectedViewWindowBeforeEdit.md)
- [ProtectedViewWindowDeactivate](Word.Application.ProtectedViewWindowDeactivate.md)
- [ProtectedViewWindowOpen](Word.Application.ProtectedViewWindowOpen.md)
- [ProtectedViewWindowSize](Word.Application.ProtectedViewWindowSize.md)
- [Quit](Word.Application.Quit(even).md)
- [WindowActivate](Word.Application.WindowActivate.md)
- [WindowBeforeDoubleClick](Word.Application.WindowBeforeDoubleClick.md)
- [WindowBeforeRightClick](Word.Application.WindowBeforeRightClick.md)
- [WindowDeactivate](Word.Application.WindowDeactivate.md)
- [WindowSelectionChange](Word.Application.WindowSelectionChange.md)
- [WindowSize](Word.Application.WindowSize.md)
- [XMLSelectionChange](Word.Application.XMLSelectionChange.md)
- [XMLValidationError](Word.Application.XMLValidationError.md)

## Methods

- [Activate](Word.Application.Activate.md)
- [AddAddress](Word.Application.AddAddress.md)
- [AutomaticChange](Word.Application.AutomaticChange.md)
- [BuildKeyCode](Word.Application.BuildKeyCode.md)
- [CentimetersToPoints](Word.Application.CentimetersToPoints.md)
- [ChangeFileOpenDirectory](Word.Application.ChangeFileOpenDirectory.md)
- [CheckGrammar](Word.Application.CheckGrammar.md)
- [CheckSpelling](Word.Application.CheckSpelling.md)
- [CleanString](Word.Application.CleanString.md)
- [CompareDocuments](Word.Application.CompareDocuments.md)
- [DDEExecute](Word.Application.DDEExecute.md)
- [DDEInitiate](Word.Application.DDEInitiate.md)
- [DDEPoke](Word.Application.DDEPoke.md)
- [DDERequest](Word.Application.DDERequest.md)
- [DDETerminate](Word.Application.DDETerminate.md)
- [DDETerminateAll](Word.Application.DDETerminateAll.md)
- [DefaultWebOptions](Word.Application.DefaultWebOptions.md)
- [GetAddress](Word.Application.GetAddress.md)
- [GetDefaultTheme](Word.Application.GetDefaultTheme.md)
- [GetSpellingSuggestions](Word.Application.GetSpellingSuggestions.md)
- [GoBack](Word.Application.GoBack.md)
- [GoForward](Word.Application.GoForward.md)
- [Help](Word.Application.Help.md)
- [HelpTool](Word.Application.HelpTool.md)
- [InchesToPoints](Word.Application.InchesToPoints.md)
- [Keyboard](Word.Application.Keyboard.md)
- [KeyboardBidi](Word.Application.KeyboardBidi.md)
- [KeyboardLatin](Word.Application.KeyboardLatin.md)
- [KeyString](Word.Application.KeyString.md)
- [LinesToPoints](Word.Application.LinesToPoints.md)
- [ListCommands](Word.Application.ListCommands.md)
- [LoadMasterList](Word.Application.LoadMasterList.md)
- [LookupNameProperties](Word.Application.LookupNameProperties.md)
- [MergeDocuments](Word.Application.MergeDocuments.md)
- [MillimetersToPoints](Word.Application.MillimetersToPoints.md)
- [Move](Word.Application.Move.md)
- [NewWindow](Word.Application.NewWindow.md)
- [NextLetter](Word.Application.Application.NextLetter.md)
- [OnTime](Word.Application.OnTime.md)
- [OrganizerCopy](Word.Application.OrganizerCopy.md)
- [OrganizerDelete](Word.Application.OrganizerDelete.md)
- [OrganizerRename](Word.Application.OrganizerRename.md)
- [PicasToPoints](Word.Application.PicasToPoints.md)
- [PixelsToPoints](Word.Application.PixelsToPoints.md)
- [PointsToCentimeters](Word.Application.PointsToCentimeters.md)
- [PointsToInches](Word.Application.PointsToInches.md)
- [PointsToLines](Word.Application.PointsToLines.md)
- [PointsToMillimeters](Word.Application.PointsToMillimeters.md)
- [PointsToPicas](Word.Application.PointsToPicas.md)
- [PointsToPixels](Word.Application.PointsToPixels.md)
- [PrintOut](Word.Application.PrintOut.md)
- [ProductCode](Word.Application.ProductCode.md)
- [PutFocusInMailHeader](Word.Application.PutFocusInMailHeader.md)
- [Quit](Word.Application.Quit(method).md)
- [Repeat](Word.Application.Repeat.md)
- [ResetIgnoreAll](Word.Application.ResetIgnoreAll.md)
- [Resize](Word.Application.Resize.md)
- [Run](Word.Application.Run.md)
- [ScreenRefresh](Word.Application.ScreenRefresh.md)
- [SetDefaultTheme](Word.Application.SetDefaultTheme.md)
- [ShowClipboard](Word.Application.ShowClipboard.md)
- [ShowMe](Word.Application.ShowMe.md)
- [SubstituteFont](Word.Application.SubstituteFont.md)
- [ToggleKeyboard](Word.Application.ToggleKeyboard.md)

## Properties

- [ActiveDocument](Word.Application.ActiveDocument.md)
- [ActiveEncryptionSession](Word.Application.ActiveEncryptionSession.md)
- [ActivePrinter](Word.Application.ActivePrinter.md)
- [ActiveProtectedViewWindow](Word.Application.ActiveProtectedViewWindow.md)
- [ActiveWindow](Word.Application.ActiveWindow.md)
- [AddIns](Word.Application.AddIns.md)
- [Application](Word.Application.Application.md)
- [ArbitraryXMLSupportAvailable](Word.Application.ArbitraryXMLSupportAvailable.md)
- [Assistance](Word.Application.Assistance.md)
- [AutoCaptions](Word.Application.AutoCaptions.md)
- [AutoCorrect](Word.Application.AutoCorrect.md)
- [AutoCorrectEmail](Word.Application.AutoCorrectEmail.md)
- [AutomationSecurity](Word.Application.AutomationSecurity.md)
- [BackgroundPrintingStatus](Word.Application.BackgroundPrintingStatus.md)
- [BackgroundSavingStatus](Word.Application.BackgroundSavingStatus.md)
- [Bibliography](Word.Application.Bibliography.md)
- [BrowseExtraFileTypes](Word.Application.BrowseExtraFileTypes.md)
- [Browser](Word.Application.Browser.md)
- [Build](Word.Application.Build.md)
- [CapsLock](Word.Application.CapsLock.md)
- [Caption](Word.Application.Caption.md)
- [CaptionLabels](Word.Application.CaptionLabels.md)
- [ChartDataPointTrack](Word.application.chartdatapointtrack.md)
- [CheckLanguage](Word.Application.CheckLanguage.md)
- [COMAddIns](Word.Application.COMAddIns.md)
- [CommandBars](Word.Application.CommandBars.md)
- [Creator](Word.Application.Creator.md)
- [CustomDictionaries](Word.Application.CustomDictionaries.md)
- [CustomizationContext](Word.Application.CustomizationContext.md)
- [DefaultLegalBlackline](Word.Application.DefaultLegalBlackline.md)
- [DefaultSaveFormat](Word.Application.DefaultSaveFormat.md)
- [DefaultTableSeparator](Word.Application.DefaultTableSeparator.md)
- [Dialogs](Word.Application.Dialogs.md)
- [DisplayAlerts](Word.Application.DisplayAlerts.md)
- [DisplayAutoCompleteTips](Word.Application.DisplayAutoCompleteTips.md)
- [DisplayDocumentInformationPanel](Word.Application.DisplayDocumentInformationPanel.md)
- [DisplayRecentFiles](Word.Application.DisplayRecentFiles.md)
- [DisplayScreenTips](Word.Application.DisplayScreenTips.md)
- [DisplayScrollBars](Word.Application.DisplayScrollBars.md)
- [Documents](Word.Application.Documents.md)
- [DontResetInsertionPointProperties](Word.Application.DontResetInsertionPointProperties.md)
- [EmailOptions](Word.Application.EmailOptions.md)
- [EmailTemplate](Word.Application.EmailTemplate.md)
- [EnableCancelKey](Word.Application.EnableCancelKey.md)
- [FeatureInstall](Word.Application.FeatureInstall.md)
- [FileConverters](Word.Application.FileConverters.md)
- [FileDialog](Word.Application.FileDialog.md)
- [FileValidation](Word.Application.FileValidation.md)
- [FindKey](Word.Application.FindKey.md)
- [FocusInMailHeader](Word.Application.FocusInMailHeader.md)
- [FontNames](Word.Application.FontNames.md)
- [HangulHanjaDictionaries](Word.Application.HangulHanjaDictionaries.md)
- [Height](Word.Application.Height.md)
- [International](Word.Application.International.md)
- [IsObjectValid](Word.Application.IsObjectValid.md)
- [IsSandboxed](Word.Application.IsSandboxed.md)
- [KeyBindings](Word.Application.KeyBindings.md)
- [KeysBoundTo](Word.Application.KeysBoundTo.md)
- [LandscapeFontNames](Word.Application.LandscapeFontNames.md)
- [Language](Word.Application.Language.md)
- [Languages](Word.Application.Languages.md)
- [LanguageSettings](Word.Application.LanguageSettings.md)
- [Left](Word.Application.Left.md)
- [ListGalleries](Word.Application.ListGalleries.md)
- [MacroContainer](Word.Application.MacroContainer.md)
- [MailingLabel](Word.Application.MailingLabel.md)
- [MailMessage](Word.Application.MailMessage.md)
- [MailSystem](Word.Application.MailSystem.md)
- [MAPIAvailable](Word.Application.MAPIAvailable.md)
- [MathCoprocessorAvailable](Word.Application.MathCoprocessorAvailable.md)
- [MouseAvailable](Word.Application.MouseAvailable.md)
- [Name](Word.Application.Name.md)
- [NewDocument](Word.Application.NewDocument(property).md)
- [NormalTemplate](Word.Application.NormalTemplate.md)
- [NumLock](Word.Application.NumLock.md)
- [OMathAutoCorrect](Word.Application.OMathAutoCorrect.md)
- [OpenAttachmentsInFullScreen](Word.Application.OpenAttachmentsInFullScreen.md)
- [Options](Word.Application.Options.md)
- [Parent](Word.Application.Parent.md)
- [Path](Word.Application.Path.md)
- [PathSeparator](Word.Application.PathSeparator.md)
- [PickerDialog](Word.Application.PickerDialog.md)
- [PortraitFontNames](Word.Application.PortraitFontNames.md)
- [PrintPreview](Word.Application.PrintPreview.md)
- [ProtectedViewWindows](Word.Application.ProtectedViewWindows.md)
- [RecentFiles](Word.Application.RecentFiles.md)
- [RestrictLinkedStyles](Word.Application.RestrictLinkedStyles.md)
- [ScreenUpdating](Word.Application.ScreenUpdating.md)
- [Selection](Word.Application.Selection.md)
- [ShowAnimation](Word.application.showanimation.md)
- [ShowStartupDialog](Word.Application.ShowStartupDialog.md)
- [ShowStylePreviews](Word.Application.ShowStylePreviews.md)
- [ShowVisualBasicEditor](Word.Application.ShowVisualBasicEditor.md)
- [SmartArtColors](Word.Application.SmartArtColors.md)
- [SmartArtLayouts](Word.Application.SmartArtLayouts.md)
- [SmartArtQuickStyles](Word.Application.SmartArtQuickStyles.md)
- [SpecialMode](Word.Application.SpecialMode.md)
- [StartupPath](Word.Application.StartupPath.md)
- [StatusBar](Word.Application.StatusBar.md)
- [SynonymInfo](Word.Application.SynonymInfo.md)
- [System](Word.Application.System.md)
- [TaskPanes](Word.Application.TaskPanes.md)
- [Tasks](Word.Application.Tasks.md)
- [Templates](Word.Application.Templates.md)
- [Top](Word.Application.Top.md)
- [UndoRecord](Word.Application.UndoRecord.md)
- [UsableHeight](Word.Application.UsableHeight.md)
- [UsableWidth](Word.Application.UsableWidth.md)
- [UserAddress](Word.Application.UserAddress.md)
- [UserControl](Word.Application.UserControl.md)
- [UserInitials](Word.Application.UserInitials.md)
- [UserName](Word.Application.UserName.md)
- [VBE](Word.Application.VBE.md)
- [Version](Word.Application.Version.md)
- [Visible](Word.Application.Visible.md)
- [Width](Word.Application.Width.md)
- [Windows](Word.Application.Windows.md)
- [WindowState](Word.Application.WindowState.md)
- [WordBasic](Word.Application.WordBasic.md)
- [XMLNamespaces](Word.Application.XMLNamespaces.md)


## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
