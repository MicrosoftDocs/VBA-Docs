---
title: Document Object (Word)
keywords: vbawd10.chm2411
f1_keywords:
- vbawd10.chm2411
ms.prod: word
api_name:
- Word.Document
ms.assetid: 8d83487a-2345-a036-a916-971c9db5b7fb
ms.date: 06/08/2017
---


# Document Object (Word)

Represents a document. The  **Document** object is a member of the **[Documents](https://msdn.microsoft.com/en-us/vba/word-vba/articles/documents-object-word)** collection. The **Documents** collection contains all the **Document** objects that are currently open in Word.


## Remarks

Use  **Documents** (Index), where Index is the document name or index number, to return a single **Document** object. The following example closes the document named "Report.doc" without saving changes.


```
Documents("Report.doc").Close SaveChanges:=wdDoNotSaveChanges
```

The index number represents the position of the document in the  **Documents** collection. The following example activates the first document in the **Documents** collection.




```
Documents(1).Activate
```

### Using ActiveDocument

You can use the  **[ActiveDocument](https://msdn.microsoft.com/en-us/vba/word-vba/articles/application-activedocument-property-word)** property to refer to the document with the focus. The following example uses the **[Activate](https://msdn.microsoft.com/en-us/vba/word-vba/articles/document-activate-method-word)** method to activate the document named "Document 1." The example also sets the page orientation to landscape mode and then prints the document.




```
Documents("Document1").Activate 
ActiveDocument.PageSetup.Orientation = wdOrientLandscape 
ActiveDocument.PrintOut
```


## Members


### Events



|**Name**|
|:-----|
|[BuildingBlockInsert](../../../api/Word.Document.BuildingBlockInsert.md)|
|[Close](../../../api/Word.Document.Close(even).md)|
|[ContentControlAfterAdd](../../../api/Word.Document.ContentControlAfterAdd.md)|
|[ContentControlBeforeContentUpdate](../../../api/Word.Document.ContentControlBeforeContentUpdate.md)|
|[ContentControlBeforeDelete](../../../api/Word.Document.ContentControlBeforeDelete.md)|
|[ContentControlBeforeStoreUpdate](../../../api/Word.Document.ContentControlBeforeStoreUpdate.md)|
|[ContentControlOnEnter](../../../api/Word.Document.ContentControlOnEnter.md)|
|[ContentControlOnExit](../../../api/Word.Document.ContentControlOnExit.md)|
|[New](../../../api/Word.Document.New.md)|
|[Open](../../../api/Word.Document.Open.md)|
|[Sync](../../../api/Word.Document.Sync(even).md)|
|[XMLAfterInsert](../../../api/Word.Document.XMLAfterInsert.md)|
|[XMLBeforeDelete](../../../api/Word.Document.XMLBeforeDelete.md)|

### Methods



|**Name**|
|:-----|
|[AcceptAllRevisions](../../../api/Word.Document.AcceptAllRevisions.md)|
|[AcceptAllRevisionsShown](../../../api/Word.Document.AcceptAllRevisionsShown.md)|
|[Activate](../../../api/Word.Document.Activate.md)|
|[AddToFavorites](../../../api/Word.Document.AddToFavorites.md)|
|[ApplyDocumentTheme](../../../api/Word.ApplyDocumentTheme.md)|
|[ApplyQuickStyleSet2](../../../api/Word.Document.ApplyQuickStyleSet2.md)|
|[ApplyTheme](../../../api/Word.Document.ApplyTheme.md)|
|[AutoFormat](../../../api/Word.Document.AutoFormat.md)|
|[CanCheckin](../../../api/Word.Document.CanCheckin.md)|
|[CheckConsistency](../../../api/Word.Document.CheckConsistency.md)|
|[CheckGrammar](../../../api/Word.Document.CheckGrammar.md)|
|[CheckIn](../../../api/Word.Document.CheckIn.md)|
|[CheckInWithVersion](../../../api/Word.Document.CheckInWithVersion.md)|
|[CheckSpelling](../../../api/Word.Document.CheckSpelling.md)|
|[Close](../../../api/Word.Document.Close(method).md)|
|[ClosePrintPreview](../../../api/Word.Document.ClosePrintPreview.md)|
|[Compare](../../../api/Word.Document.Compare.md)|
|[ComputeStatistics](../../../api/Word.Document.ComputeStatistics.md)|
|[Convert](../../../api/Word.Document.Convert.md)|
|[ConvertAutoHyphens](../../../api/Word.Document.ConvertAutoHyphens.md)|
|[ConvertNumbersToText](../../../api/Word.Document.ConvertNumbersToText.md)|
|[ConvertVietDoc](../../../api/Word.Document.ConvertVietDoc.md)|
|[CopyStylesFromTemplate](../../../api/Word.Document.CopyStylesFromTemplate.md)|
|[CountNumberedItems](../../../api/Word.Document.CountNumberedItems.md)|
|[CreateLetterContent](../../../api/Word.Document.CreateLetterContent.md)|
|[DataForm](../../../api/Word.Document.DataForm.md)|
|[DeleteAllComments](../../../api/Word.Document.DeleteAllComments.md)|
|[DeleteAllCommentsShown](../../../api/Word.Document.DeleteAllCommentsShown.md)|
|[DeleteAllEditableRanges](../../../api/Word.Document.DeleteAllEditableRanges.md)|
|[DeleteAllInkAnnotations](../../../api/Word.Document.DeleteAllInkAnnotations.md)|
|[DetectLanguage](../../../api/Word.Document.DetectLanguage.md)|
|[DowngradeDocument](../../../api/Word.Document.DowngradeDocument.md)|
|[EndReview](../../../api/Word.Document.EndReview.md)|
|[ExportAsFixedFormat](../../../api/Word.Document.ExportAsFixedFormat.md)|
|[FitToPages](../../../api/Word.Document.FitToPages.md)|
|[FollowHyperlink](../../../api/Word.Document.FollowHyperlink.md)|
|[FreezeLayout](../../../api/Word.Document.FreezeLayout.md)|
|[GetCrossReferenceItems](../../../api/Word.Document.GetCrossReferenceItems.md)|
|[GetLetterContent](../../../api/Word.Document.GetLetterContent.md)|
|[GetWorkflowTasks](../../../api/Word.Document.GetWorkflowTasks.md)|
|[GetWorkflowTemplates](../../../api/Word.Document.GetWorkflowTemplates.md)|
|[GoTo](../../../api/Word.Document.GoTo.md)|
|[LockServerFile](../../../api/Word.Document.LockServerFile.md)|
|[MakeCompatibilityDefault](../../../api/Word.Document.MakeCompatibilityDefault.md)|
|[ManualHyphenation](../../../api/Word.Document.ManualHyphenation.md)|
|[Merge](../../../api/Word.Document.Merge.md)|
|[Post](../../../api/Word.Document.Post.md)|
|[PresentIt](../../../api/Word.Document.PresentIt.md)|
|[PrintOut](../../../api/Word.Document.PrintOut.md)|
|[PrintPreview](../../../api/Word.Document.PrintPreview.md)|
|[Protect](../../../api/Word.document.protect.md)|
|[Range](../../../api/Word.Document.Range.md)|
|[Redo](../../../api/Word.Document.Redo.md)|
|[RejectAllRevisions](../../../api/Word.Document.RejectAllRevisions.md)|
|[RejectAllRevisionsShown](../../../api/Word.Document.RejectAllRevisionsShown.md)|
|[Reload](../../../api/Word.Document.Reload.md)|
|[ReloadAs](../../../api/Word.Document.ReloadAs.md)|
|[RemoveDocumentInformation](../../../api/Word.Document.RemoveDocumentInformation.md)|
|[RemoveLockedStyles](../../../api/Word.Document.RemoveLockedStyles.md)|
|[RemoveNumbers](../../../api/Word.Document.RemoveNumbers.md)|
|[RemoveTheme](../../../api/Word.Document.RemoveTheme.md)|
|[Repaginate](../../../api/Word.Document.Repaginate.md)|
|[ReplyWithChanges](../../../api/Word.Document.ReplyWithChanges.md)|
|[ResetFormFields](../../../api/Word.Document.ResetFormFields.md)|
|[ReturnToLastReadPosition](../../../api/Word.document.returntolastreadposition.md)|
|[RunAutoMacro](../../../api/Word.Document.RunAutoMacro.md)|
|[RunLetterWizard](../../../api/Word.Document.RunLetterWizard.md)|
|[Save](../../../api/Word.Document.Save.md)|
|[SaveAs2](../../../api/Word.SaveAs2.md)|
|[SaveAsQuickStyleSet](../../../api/Word.Document.SaveAsQuickStyleSet.md)|
|[Select](../../../api/Word.Document.Select.md)|
|[SelectAllEditableRanges](../../../api/Word.Document.SelectAllEditableRanges.md)|
|[SelectContentControlsByTag](../../../api/Word.Document.SelectContentControlsByTag.md)|
|[SelectContentControlsByTitle](../../../api/Word.Document.SelectContentControlsByTitle.md)|
|[SelectLinkedControls](../../../api/Word.Document.SelectLinkedControls.md)|
|[SelectNodes](../../../api/Word.Document.SelectNodes.md)|
|[SelectSingleNode](../../../api/Word.Document.SelectSingleNode.md)|
|[SelectUnlinkedControls](../../../api/Word.Document.SelectUnlinkedControls.md)|
|[SendFax](../../../api/Word.Document.SendFax.md)|
|[SendFaxOverInternet](../../../api/Word.Document.SendFaxOverInternet.md)|
|[SendForReview](../../../api/Word.Document.SendForReview.md)|
|[SendMail](../../../api/Word.Document.SendMail.md)|
|[SetCompatibilityMode](../../../api/Word.SetCompatibilityMode.md)|
|[SetDefaultTableStyle](../../../api/Word.Document.SetDefaultTableStyle.md)|
|[SetLetterContent](../../../api/Word.Document.SetLetterContent.md)|
|[SetPasswordEncryptionOptions](../../../api/Word.Document.SetPasswordEncryptionOptions.md)|
|[ToggleFormsDesign](../../../api/Word.Document.ToggleFormsDesign.md)|
|[TransformDocument](../../../api/Word.Document.TransformDocument.md)|
|[Undo](../../../api/Word.Document.Undo.md)|
|[UndoClear](../../../api/Word.Document.UndoClear.md)|
|[Unprotect](../../../api/Word.Document.Unprotect.md)|
|[UpdateStyles](../../../api/Word.Document.UpdateStyles.md)|
|[ViewCode](../../../api/Word.Document.ViewCode.md)|
|[ViewPropertyBrowser](../../../api/Word.Document.ViewPropertyBrowser.md)|
|[WebPagePreview](../../../api/Word.Document.WebPagePreview.md)|

### Properties



|**Name**|
|:-----|
|[ActiveTheme](../../../api/Word.Document.ActiveTheme.md)|
|[ActiveThemeDisplayName](../../../api/Word.Document.ActiveThemeDisplayName.md)|
|[ActiveWindow](../../../api/Word.Document.ActiveWindow.md)|
|[ActiveWritingStyle](../../../api/Word.Document.ActiveWritingStyle.md)|
|[Application](../../../api/Word.Document.Application.md)|
|[AttachedTemplate](../../../api/Word.Document.AttachedTemplate.md)|
|[AutoFormatOverride](../../../api/Word.Document.AutoFormatOverride.md)|
|[AutoHyphenation](../../../api/Word.Document.AutoHyphenation.md)|
|[Background](../../../api/Word.Document.Background.md)|
|[Bibliography](../../../api/Word.Document.Bibliography.md)|
|[Bookmarks](../../../api/Word.Document.Bookmarks.md)|
|[Broadcast](../../../api/Word.document.broadcast.md)|
|[BuiltInDocumentProperties](../../../api/Word.Document.BuiltInDocumentProperties.md)|
|[Characters](../../../api/Word.Document.Characters.md)|
|[ChartDataPointTrack](../../../api/Word.document.chartdatapointtrack.md)|
|[ClickAndTypeParagraphStyle](../../../api/Word.Document.ClickAndTypeParagraphStyle.md)|
|[CoAuthoring](../../../api/Word.Document.CoAuthoring.md)|
|[CodeName](../../../api/Word.Document.CodeName.md)|
|[CommandBars](../../../api/Word.Document.CommandBars.md)|
|[Comments](../../../api/Word.Document.Comments.md)|
|[Compatibility](../../../api/Word.Document.Compatibility.md)|
|[CompatibilityMode](../../../api/Word.Document.CompatibilityMode.md)|
|[ConsecutiveHyphensLimit](../../../api/Word.Document.ConsecutiveHyphensLimit.md)|
|[Container](../../../api/Word.Document.Container.md)|
|[Content](../../../api/Word.Document.Content.md)|
|[ContentControls](../../../api/Word.Document.ContentControls.md)|
|[ContentTypeProperties](../../../api/Word.Document.ContentTypeProperties.md)|
|[Creator](../../../api/Word.Document.Creator.md)|
|[CurrentRsid](../../../api/Word.Document.CurrentRsid.md)|
|[CustomDocumentProperties](../../../api/Word.Document.CustomDocumentProperties.md)|
|[CustomXMLParts](../../../api/Word.Document.CustomXMLParts.md)|
|[DefaultTableStyle](../../../api/Word.Document.DefaultTableStyle.md)|
|[DefaultTabStop](../../../api/Word.Document.DefaultTabStop.md)|
|[DefaultTargetFrame](../../../api/Word.Document.DefaultTargetFrame.md)|
|[DisableFeatures](../../../api/Word.Document.DisableFeatures.md)|
|[DisableFeaturesIntroducedAfter](../../../api/Word.Document.DisableFeaturesIntroducedAfter.md)|
|[DocumentInspectors](../../../api/Word.Document.DocumentInspectors.md)|
|[DocumentLibraryVersions](../../../api/Word.Document.DocumentLibraryVersions.md)|
|[DocumentTheme](../../../api/Word.Document.DocumentTheme.md)|
|[DoNotEmbedSystemFonts](../../../api/Word.Document.DoNotEmbedSystemFonts.md)|
|[Email](../../../api/Word.Document.Email.md)|
|[EmbedLinguisticData](../../../api/Word.Document.EmbedLinguisticData.md)|
|[EmbedTrueTypeFonts](../../../api/Word.Document.EmbedTrueTypeFonts.md)|
|[EncryptionProvider](../../../api/Word.Document.EncryptionProvider.md)|
|[Endnotes](../../../api/Word.Document.Endnotes.md)|
|[EnforceStyle](../../../api/Word.Document.EnforceStyle.md)|
|[Envelope](../../../api/Word.Document.Envelope.md)|
|[FarEastLineBreakLanguage](../../../api/Word.Document.FarEastLineBreakLanguage.md)|
|[FarEastLineBreakLevel](../../../api/Word.Document.FarEastLineBreakLevel.md)|
|[Fields](../../../api/Word.Document.Fields.md)|
|[Final](../../../api/Word.Document.Final.md)|
|[Footnotes](../../../api/Word.Document.Footnotes.md)|
|[FormattingShowClear](../../../api/Word.Document.FormattingShowClear.md)|
|[FormattingShowFilter](../../../api/Word.Document.FormattingShowFilter.md)|
|[FormattingShowFont](../../../api/Word.Document.FormattingShowFont.md)|
|[FormattingShowNextLevel](../../../api/Word.Document.FormattingShowNextLevel.md)|
|[FormattingShowNumbering](../../../api/Word.Document.FormattingShowNumbering.md)|
|[FormattingShowParagraph](../../../api/Word.Document.FormattingShowParagraph.md)|
|[FormattingShowUserStyleName](../../../api/Word.Document.FormattingShowUserStyleName.md)|
|[FormFields](../../../api/Word.Document.FormFields.md)|
|[FormsDesign](../../../api/Word.Document.FormsDesign.md)|
|[Frames](../../../api/Word.Document.Frames.md)|
|[Frameset](../../../api/Word.Document.Frameset.md)|
|[FullName](../../../api/Word.Document.FullName.md)|
|[GrammarChecked](../../../api/Word.Document.GrammarChecked.md)|
|[GrammaticalErrors](../../../api/Word.Document.GrammaticalErrors.md)|
|[GridDistanceHorizontal](../../../api/Word.Document.GridDistanceHorizontal.md)|
|[GridDistanceVertical](../../../api/Word.Document.GridDistanceVertical.md)|
|[GridOriginFromMargin](../../../api/Word.Document.GridOriginFromMargin.md)|
|[GridOriginHorizontal](../../../api/Word.Document.GridOriginHorizontal.md)|
|[GridOriginVertical](../../../api/Word.Document.GridOriginVertical.md)|
|[GridSpaceBetweenHorizontalLines](../../../api/Word.Document.GridSpaceBetweenHorizontalLines.md)|
|[GridSpaceBetweenVerticalLines](../../../api/Word.Document.GridSpaceBetweenVerticalLines.md)|
|[HasPassword](../../../api/Word.Document.HasPassword.md)|
|[HasVBProject](../../../api/Word.Document.HasVBProject.md)|
|[HTMLDivisions](../../../api/Word.Document.HTMLDivisions.md)|
|[Hyperlinks](../../../api/Word.Document.Hyperlinks.md)|
|[HyphenateCaps](../../../api/Word.Document.HyphenateCaps.md)|
|[HyphenationZone](../../../api/Word.Document.HyphenationZone.md)|
|[Indexes](../../../api/Word.Document.Indexes.md)|
|[InlineShapes](../../../api/Word.Document.InlineShapes.md)|
|[IsInAutosave](../../../api/Word.document.isinautosave.md)|
|[IsMasterDocument](../../../api/Word.Document.IsMasterDocument.md)|
|[IsSubdocument](../../../api/Word.Document.IsSubdocument.md)|
|[JustificationMode](../../../api/Word.Document.JustificationMode.md)|
|[KerningByAlgorithm](../../../api/Word.Document.KerningByAlgorithm.md)|
|[Kind](../../../api/Word.Document.Kind.md)|
|[LanguageDetected](../../../api/Word.Document.LanguageDetected.md)|
|[ListParagraphs](../../../api/Word.Document.ListParagraphs.md)|
|[Lists](../../../api/Word.Document.Lists.md)|
|[ListTemplates](../../../api/Word.Document.ListTemplates.md)|
|[LockQuickStyleSet](../../../api/Word.Document.LockQuickStyleSet.md)|
|[LockTheme](../../../api/Word.Document.LockTheme.md)|
|[MailEnvelope](../../../api/Word.Document.MailEnvelope.md)|
|[MailMerge](../../../api/Word.Document.MailMerge.md)|
|[Name](../../../api/Word.Document.Name.md)|
|[NoLineBreakAfter](../../../api/Word.Document.NoLineBreakAfter.md)|
|[NoLineBreakBefore](../../../api/Word.Document.NoLineBreakBefore.md)|
|[OMathBreakBin](../../../api/Word.Document.OMathBreakBin.md)|
|[OMathBreakSub](../../../api/Word.Document.OMathBreakSub.md)|
|[OMathFontName](../../../api/Word.Document.OMathFontName.md)|
|[OMathIntSubSupLim](../../../api/Word.Document.OMathIntSubSupLim.md)|
|[OMathJc](../../../api/Word.Document.OMathJc.md)|
|[OMathLeftMargin](../../../api/Word.Document.OMathLeftMargin.md)|
|[OMathNarySupSubLim](../../../api/Word.Document.OMathNarySupSubLim.md)|
|[OMathRightMargin](../../../api/Word.Document.OMathRightMargin.md)|
|[OMaths](../../../api/Word.Document.OMaths.md)|
|[OMathSmallFrac](../../../api/Word.Document.OMathSmallFrac.md)|
|[OMathWrap](../../../api/Word.Document.OMathWrap.md)|
|[OpenEncoding](../../../api/Word.Document.OpenEncoding.md)|
|[OptimizeForWord97](../../../api/Word.Document.OptimizeForWord97.md)|
|[OriginalDocumentTitle](../../../api/Word.Document.OriginalDocumentTitle.md)|
|[PageSetup](../../../api/Word.Document.PageSetup.md)|
|[Paragraphs](../../../api/Word.Document.Paragraphs.md)|
|[Parent](../../../api/Word.Document.Parent.md)|
|[Password](../../../api/Word.Document.Password.md)|
|[PasswordEncryptionAlgorithm](../../../api/Word.Document.PasswordEncryptionAlgorithm.md)|
|[PasswordEncryptionFileProperties](../../../api/Word.Document.PasswordEncryptionFileProperties.md)|
|[PasswordEncryptionKeyLength](../../../api/Word.Document.PasswordEncryptionKeyLength.md)|
|[PasswordEncryptionProvider](../../../api/Word.Document.PasswordEncryptionProvider.md)|
|[Path](../../../api/Word.Document.Path.md)|
|[Permission](../../../api/Word.Document.Permission.md)|
|[PrintFormsData](../../../api/Word.Document.PrintFormsData.md)|
|[PrintPostScriptOverText](../../../api/Word.Document.PrintPostScriptOverText.md)|
|[PrintRevisions](../../../api/Word.Document.PrintRevisions.md)|
|[ProtectionType](../../../api/Word.Document.ProtectionType.md)|
|[ReadabilityStatistics](../../../api/Word.Document.ReadabilityStatistics.md)|
|[ReadingLayoutSizeX](../../../api/Word.Document.ReadingLayoutSizeX.md)|
|[ReadingLayoutSizeY](../../../api/Word.Document.ReadingLayoutSizeY.md)|
|[ReadingModeLayoutFrozen](../../../api/Word.Document.ReadingModeLayoutFrozen.md)|
|[ReadOnly](../../../api/Word.Document.ReadOnly.md)|
|[ReadOnlyRecommended](../../../api/Word.Document.ReadOnlyRecommended.md)|
|[RemoveDateAndTime](../../../api/Word.Document.RemoveDateAndTime.md)|
|[RemovePersonalInformation](../../../api/Word.Document.RemovePersonalInformation.md)|
|[Research](../../../api/Word.Document.Research.md)|
|[RevisedDocumentTitle](../../../api/Word.Document.RevisedDocumentTitle.md)|
|[Revisions](../../../api/Word.Document.Revisions.md)|
|[Saved](../../../api/Word.Document.Saved.md)|
|[SaveEncoding](../../../api/Word.Document.SaveEncoding.md)|
|[SaveFormat](../../../api/Word.Document.SaveFormat.md)|
|[SaveFormsData](../../../api/Word.Document.SaveFormsData.md)|
|[SaveSubsetFonts](../../../api/Word.Document.SaveSubsetFonts.md)|
|[Scripts](../../../api/Word.Document.Scripts.md)|
|[Sections](../../../api/Word.Document.Sections.md)|
|[Sentences](../../../api/Word.Document.Sentences.md)|
|[ServerPolicy](../../../api/Word.Document.ServerPolicy.md)|
|[Shapes](../../../api/Word.Document.Shapes.md)|
|[ShowGrammaticalErrors](../../../api/Word.Document.ShowGrammaticalErrors.md)|
|[ShowSpellingErrors](../../../api/Word.Document.ShowSpellingErrors.md)|
|[Signatures](../../../api/Word.Document.Signatures.md)|
|[SmartDocument](../../../api/Word.Document.SmartDocument.md)|
|[SnapToGrid](../../../api/Word.Document.SnapToGrid.md)|
|[SnapToShapes](../../../api/Word.Document.SnapToShapes.md)|
|[SpellingChecked](../../../api/Word.Document.SpellingChecked.md)|
|[SpellingErrors](../../../api/Word.Document.SpellingErrors.md)|
|[StoryRanges](../../../api/Word.Document.StoryRanges.md)|
|[Styles](../../../api/Word.Document.Styles.md)|
|[StyleSheets](../../../api/Word.Document.StyleSheets.md)|
|[StyleSortMethod](../../../api/Word.Document.StyleSortMethod.md)|
|[Subdocuments](../../../api/Word.Document.Subdocuments.md)|
|[Sync](../../../api/Word.Document.Sync(property).md)|
|[Tables](../../../api/Word.Document.Tables.md)|
|[TablesOfAuthorities](../../../api/Word.Document.TablesOfAuthorities.md)|
|[TablesOfAuthoritiesCategories](../../../api/Word.Document.TablesOfAuthoritiesCategories.md)|
|[TablesOfContents](../../../api/Word.Document.TablesOfContents.md)|
|[TablesOfFigures](../../../api/Word.Document.TablesOfFigures.md)|
|[TextEncoding](../../../api/Word.Document.TextEncoding.md)|
|[TextLineEnding](../../../api/Word.Document.TextLineEnding.md)|
|[TrackFormatting](../../../api/Word.Document.TrackFormatting.md)|
|[TrackMoves](../../../api/Word.Document.TrackMoves.md)|
|[TrackRevisions](../../../api/Word.Document.TrackRevisions.md)|
|[Type](../../../api/Word.Document.Type.md)|
|[UpdateStylesOnOpen](../../../api/Word.Document.UpdateStylesOnOpen.md)|
|[UseMathDefaults](../../../api/Word.Document.UseMathDefaults.md)|
|[UserControl](../../../api/Word.Document.UserControl.md)|
|[Variables](../../../api/Word.Document.Variables.md)|
|[VBASigned](../../../api/Word.Document.VBASigned.md)|
|[VBProject](../../../api/Word.Document.VBProject.md)|
|[WebOptions](../../../api/Word.Document.WebOptions.md)|
|[Windows](../../../api/Word.Document.Windows.md)|
|[WordOpenXML](../../../api/Word.Document.WordOpenXML.md)|
|[Words](../../../api/Word.Document.Words.md)|
|[WritePassword](../../../api/Word.Document.WritePassword.md)|
|[WriteReserved](../../../api/Word.Document.WriteReserved.md)|
|[XMLSaveThroughXSLT](../../../api/Word.Document.XMLSaveThroughXSLT.md)|
|[XMLSchemaReferences](../../../api/Word.Document.XMLSchemaReferences.md)|
|[XMLShowAdvancedErrors](../../../api/Word.Document.XMLShowAdvancedErrors.md)|
|[XMLUseXSLTWhenSaving](../../../api/Word.Document.XMLUseXSLTWhenSaving.md)|

## See also


#### Other resources


[Word Object Model Reference](../../../api/overview/object-model-word-vba-reference.md)

