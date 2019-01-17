---
title: Document object (Word)
keywords: vbawd10.chm2411
f1_keywords:
- vbawd10.chm2411
ms.prod: word
api_name:
- Word.Document
ms.assetid: 8d83487a-2345-a036-a916-971c9db5b7fb
ms.date: 06/08/2017
localization_priority: Priority
---


# Document object (Word)

Represents a document. The  **Document** object is a member of the **[Documents](https://msdn.microsoft.com/vba/word-vba/articles/documents-object-word)** collection. The **Documents** collection contains all the **Document** objects that are currently open in Word.


## Remarks

Use  **Documents** (Index), where Index is the document name or index number, to return a single **Document** object. The following example closes the document named "Report.doc" without saving changes.


```vb
Documents("Report.doc").Close SaveChanges:=wdDoNotSaveChanges
```

The index number represents the position of the document in the  **Documents** collection. The following example activates the first document in the **Documents** collection.




```vb
Documents(1).Activate
```

### Using ActiveDocument

You can use the  **[ActiveDocument](https://msdn.microsoft.com/vba/word-vba/articles/application-activedocument-property-word)** property to refer to the document with the focus. The following example uses the **[Activate](https://msdn.microsoft.com/vba/word-vba/articles/document-activate-method-word)** method to activate the document named "Document 1." The example also sets the page orientation to landscape mode and then prints the document.




```vb
Documents("Document1").Activate 
ActiveDocument.PageSetup.Orientation = wdOrientLandscape 
ActiveDocument.PrintOut
```


## Members


### Events



|Name|
|:-----|
|[BuildingBlockInsert](./Word.Document.BuildingBlockInsert.md)|
|[Close](./Word.Document.Close(even).md)|
|[ContentControlAfterAdd](./Word.Document.ContentControlAfterAdd.md)|
|[ContentControlBeforeContentUpdate](./Word.Document.ContentControlBeforeContentUpdate.md)|
|[ContentControlBeforeDelete](./Word.Document.ContentControlBeforeDelete.md)|
|[ContentControlBeforeStoreUpdate](./Word.Document.ContentControlBeforeStoreUpdate.md)|
|[ContentControlOnEnter](./Word.Document.ContentControlOnEnter.md)|
|[ContentControlOnExit](./Word.Document.ContentControlOnExit.md)|
|[New](./Word.Document.New.md)|
|[Open](./Word.Document.Open.md)|
|[Sync](./Word.Document.Sync(even).md)|
|[XMLAfterInsert](./Word.Document.XMLAfterInsert.md)|
|[XMLBeforeDelete](./Word.Document.XMLBeforeDelete.md)|

### Methods



|Name|
|:-----|
|[AcceptAllRevisions](./Word.Document.AcceptAllRevisions.md)|
|[AcceptAllRevisionsShown](./Word.Document.AcceptAllRevisionsShown.md)|
|[Activate](./Word.Document.Activate.md)|
|[AddToFavorites](./Word.Document.AddToFavorites.md)|
|[ApplyDocumentTheme](./Word.ApplyDocumentTheme.md)|
|[ApplyQuickStyleSet2](./Word.Document.ApplyQuickStyleSet2.md)|
|[ApplyTheme](./Word.Document.ApplyTheme.md)|
|[AutoFormat](./Word.Document.AutoFormat.md)|
|[CanCheckin](./Word.Document.CanCheckin.md)|
|[CheckConsistency](./Word.Document.CheckConsistency.md)|
|[CheckGrammar](./Word.Document.CheckGrammar.md)|
|[CheckIn](./Word.Document.CheckIn.md)|
|[CheckInWithVersion](./Word.Document.CheckInWithVersion.md)|
|[CheckSpelling](./Word.Document.CheckSpelling.md)|
|[Close](./Word.Document.Close(method).md)|
|[ClosePrintPreview](./Word.Document.ClosePrintPreview.md)|
|[Compare](./Word.Document.Compare.md)|
|[ComputeStatistics](./Word.Document.ComputeStatistics.md)|
|[Convert](./Word.Document.Convert.md)|
|[ConvertAutoHyphens](./Word.Document.ConvertAutoHyphens.md)|
|[ConvertNumbersToText](./Word.Document.ConvertNumbersToText.md)|
|[ConvertVietDoc](./Word.Document.ConvertVietDoc.md)|
|[CopyStylesFromTemplate](./Word.Document.CopyStylesFromTemplate.md)|
|[CountNumberedItems](./Word.Document.CountNumberedItems.md)|
|[CreateLetterContent](./Word.Document.CreateLetterContent.md)|
|[DataForm](./Word.Document.DataForm.md)|
|[DeleteAllComments](./Word.Document.DeleteAllComments.md)|
|[DeleteAllCommentsShown](./Word.Document.DeleteAllCommentsShown.md)|
|[DeleteAllEditableRanges](./Word.Document.DeleteAllEditableRanges.md)|
|[DeleteAllInkAnnotations](./Word.Document.DeleteAllInkAnnotations.md)|
|[DetectLanguage](./Word.Document.DetectLanguage.md)|
|[DowngradeDocument](./Word.Document.DowngradeDocument.md)|
|[EndReview](./Word.Document.EndReview.md)|
|[ExportAsFixedFormat](./Word.Document.ExportAsFixedFormat.md)|
|[FitToPages](./Word.Document.FitToPages.md)|
|[FollowHyperlink](./Word.Document.FollowHyperlink.md)|
|[FreezeLayout](./Word.Document.FreezeLayout.md)|
|[GetCrossReferenceItems](./Word.Document.GetCrossReferenceItems.md)|
|[GetLetterContent](./Word.Document.GetLetterContent.md)|
|[GetWorkflowTasks](./Word.Document.GetWorkflowTasks.md)|
|[GetWorkflowTemplates](./Word.Document.GetWorkflowTemplates.md)|
|[GoTo](./Word.Document.GoTo.md)|
|[LockServerFile](./Word.Document.LockServerFile.md)|
|[MakeCompatibilityDefault](./Word.Document.MakeCompatibilityDefault.md)|
|[ManualHyphenation](./Word.Document.ManualHyphenation.md)|
|[Merge](./Word.Document.Merge.md)|
|[Post](./Word.Document.Post.md)|
|[PresentIt](./Word.Document.PresentIt.md)|
|[PrintOut](./Word.Document.PrintOut.md)|
|[PrintPreview](./Word.Document.PrintPreview.md)|
|[Protect](./Word.document.protect.md)|
|[Range](./Word.Document.Range.md)|
|[Redo](./Word.Document.Redo.md)|
|[RejectAllRevisions](./Word.Document.RejectAllRevisions.md)|
|[RejectAllRevisionsShown](./Word.Document.RejectAllRevisionsShown.md)|
|[Reload](./Word.Document.Reload.md)|
|[ReloadAs](./Word.Document.ReloadAs.md)|
|[RemoveDocumentInformation](./Word.Document.RemoveDocumentInformation.md)|
|[RemoveLockedStyles](./Word.Document.RemoveLockedStyles.md)|
|[RemoveNumbers](./Word.Document.RemoveNumbers.md)|
|[RemoveTheme](./Word.Document.RemoveTheme.md)|
|[Repaginate](./Word.Document.Repaginate.md)|
|[ReplyWithChanges](./Word.Document.ReplyWithChanges.md)|
|[ResetFormFields](./Word.Document.ResetFormFields.md)|
|[ReturnToLastReadPosition](./Word.document.returntolastreadposition.md)|
|[RunAutoMacro](./Word.Document.RunAutoMacro.md)|
|[RunLetterWizard](./Word.Document.RunLetterWizard.md)|
|[Save](./Word.Document.Save.md)|
|[SaveAs2](./Word.SaveAs2.md)|
|[SaveAsQuickStyleSet](./Word.Document.SaveAsQuickStyleSet.md)|
|[Select](./Word.Document.Select.md)|
|[SelectAllEditableRanges](./Word.Document.SelectAllEditableRanges.md)|
|[SelectContentControlsByTag](./Word.Document.SelectContentControlsByTag.md)|
|[SelectContentControlsByTitle](./Word.Document.SelectContentControlsByTitle.md)|
|[SelectLinkedControls](./Word.Document.SelectLinkedControls.md)|
|[SelectNodes](./Word.Document.SelectNodes.md)|
|[SelectSingleNode](./Word.Document.SelectSingleNode.md)|
|[SelectUnlinkedControls](./Word.Document.SelectUnlinkedControls.md)|
|[SendFax](./Word.Document.SendFax.md)|
|[SendFaxOverInternet](./Word.Document.SendFaxOverInternet.md)|
|[SendForReview](./Word.Document.SendForReview.md)|
|[SendMail](./Word.Document.SendMail.md)|
|[SetCompatibilityMode](./Word.SetCompatibilityMode.md)|
|[SetDefaultTableStyle](./Word.Document.SetDefaultTableStyle.md)|
|[SetLetterContent](./Word.Document.SetLetterContent.md)|
|[SetPasswordEncryptionOptions](./Word.Document.SetPasswordEncryptionOptions.md)|
|[ToggleFormsDesign](./Word.Document.ToggleFormsDesign.md)|
|[TransformDocument](./Word.Document.TransformDocument.md)|
|[Undo](./Word.Document.Undo.md)|
|[UndoClear](./Word.Document.UndoClear.md)|
|[Unprotect](./Word.Document.Unprotect.md)|
|[UpdateStyles](./Word.Document.UpdateStyles.md)|
|[ViewCode](./Word.Document.ViewCode.md)|
|[ViewPropertyBrowser](./Word.Document.ViewPropertyBrowser.md)|
|[WebPagePreview](./Word.Document.WebPagePreview.md)|

### Properties



|Name|
|:-----|
|[ActiveTheme](./Word.Document.ActiveTheme.md)|
|[ActiveThemeDisplayName](./Word.Document.ActiveThemeDisplayName.md)|
|[ActiveWindow](./Word.Document.ActiveWindow.md)|
|[ActiveWritingStyle](./Word.Document.ActiveWritingStyle.md)|
|[Application](./Word.Document.Application.md)|
|[AttachedTemplate](./Word.Document.AttachedTemplate.md)|
|[AutoFormatOverride](./Word.Document.AutoFormatOverride.md)|
|[AutoHyphenation](./Word.Document.AutoHyphenation.md)|
|[Background](./Word.Document.Background.md)|
|[Bibliography](./Word.Document.Bibliography.md)|
|[Bookmarks](./Word.Document.Bookmarks.md)|
|[Broadcast](./Word.document.broadcast.md)|
|[BuiltInDocumentProperties](./Word.Document.BuiltInDocumentProperties.md)|
|[Characters](./Word.Document.Characters.md)|
|[ChartDataPointTrack](./Word.document.chartdatapointtrack.md)|
|[ClickAndTypeParagraphStyle](./Word.Document.ClickAndTypeParagraphStyle.md)|
|[CoAuthoring](./Word.Document.CoAuthoring.md)|
|[CodeName](./Word.Document.CodeName.md)|
|[CommandBars](./Word.Document.CommandBars.md)|
|[Comments](./Word.Document.Comments.md)|
|[Compatibility](./Word.Document.Compatibility.md)|
|[CompatibilityMode](./Word.Document.CompatibilityMode.md)|
|[ConsecutiveHyphensLimit](./Word.Document.ConsecutiveHyphensLimit.md)|
|[Container](./Word.Document.Container.md)|
|[Content](./Word.Document.Content.md)|
|[ContentControls](./Word.Document.ContentControls.md)|
|[ContentTypeProperties](./Word.Document.ContentTypeProperties.md)|
|[Creator](./Word.Document.Creator.md)|
|[CurrentRsid](./Word.Document.CurrentRsid.md)|
|[CustomDocumentProperties](./Word.Document.CustomDocumentProperties.md)|
|[CustomXMLParts](./Word.Document.CustomXMLParts.md)|
|[DefaultTableStyle](./Word.Document.DefaultTableStyle.md)|
|[DefaultTabStop](./Word.Document.DefaultTabStop.md)|
|[DefaultTargetFrame](./Word.Document.DefaultTargetFrame.md)|
|[DisableFeatures](./Word.Document.DisableFeatures.md)|
|[DisableFeaturesIntroducedAfter](./Word.Document.DisableFeaturesIntroducedAfter.md)|
|[DocumentInspectors](./Word.Document.DocumentInspectors.md)|
|[DocumentLibraryVersions](./Word.Document.DocumentLibraryVersions.md)|
|[DocumentTheme](./Word.Document.DocumentTheme.md)|
|[DoNotEmbedSystemFonts](./Word.Document.DoNotEmbedSystemFonts.md)|
|[Email](./Word.Document.Email.md)|
|[EmbedLinguisticData](./Word.Document.EmbedLinguisticData.md)|
|[EmbedTrueTypeFonts](./Word.Document.EmbedTrueTypeFonts.md)|
|[EncryptionProvider](./Word.Document.EncryptionProvider.md)|
|[Endnotes](./Word.Document.Endnotes.md)|
|[EnforceStyle](./Word.Document.EnforceStyle.md)|
|[Envelope](./Word.Document.Envelope.md)|
|[FarEastLineBreakLanguage](./Word.Document.FarEastLineBreakLanguage.md)|
|[FarEastLineBreakLevel](./Word.Document.FarEastLineBreakLevel.md)|
|[Fields](./Word.Document.Fields.md)|
|[Final](./Word.Document.Final.md)|
|[Footnotes](./Word.Document.Footnotes.md)|
|[FormattingShowClear](./Word.Document.FormattingShowClear.md)|
|[FormattingShowFilter](./Word.Document.FormattingShowFilter.md)|
|[FormattingShowFont](./Word.Document.FormattingShowFont.md)|
|[FormattingShowNextLevel](./Word.Document.FormattingShowNextLevel.md)|
|[FormattingShowNumbering](./Word.Document.FormattingShowNumbering.md)|
|[FormattingShowParagraph](./Word.Document.FormattingShowParagraph.md)|
|[FormattingShowUserStyleName](./Word.Document.FormattingShowUserStyleName.md)|
|[FormFields](./Word.Document.FormFields.md)|
|[FormsDesign](./Word.Document.FormsDesign.md)|
|[Frames](./Word.Document.Frames.md)|
|[Frameset](./Word.Document.Frameset.md)|
|[FullName](./Word.Document.FullName.md)|
|[GrammarChecked](./Word.Document.GrammarChecked.md)|
|[GrammaticalErrors](./Word.Document.GrammaticalErrors.md)|
|[GridDistanceHorizontal](./Word.Document.GridDistanceHorizontal.md)|
|[GridDistanceVertical](./Word.Document.GridDistanceVertical.md)|
|[GridOriginFromMargin](./Word.Document.GridOriginFromMargin.md)|
|[GridOriginHorizontal](./Word.Document.GridOriginHorizontal.md)|
|[GridOriginVertical](./Word.Document.GridOriginVertical.md)|
|[GridSpaceBetweenHorizontalLines](./Word.Document.GridSpaceBetweenHorizontalLines.md)|
|[GridSpaceBetweenVerticalLines](./Word.Document.GridSpaceBetweenVerticalLines.md)|
|[HasPassword](./Word.Document.HasPassword.md)|
|[HasVBProject](./Word.Document.HasVBProject.md)|
|[HTMLDivisions](./Word.Document.HTMLDivisions.md)|
|[Hyperlinks](./Word.Document.Hyperlinks.md)|
|[HyphenateCaps](./Word.Document.HyphenateCaps.md)|
|[HyphenationZone](./Word.Document.HyphenationZone.md)|
|[Indexes](./Word.Document.Indexes.md)|
|[InlineShapes](./Word.Document.InlineShapes.md)|
|[IsInAutosave](./Word.document.isinautosave.md)|
|[IsMasterDocument](./Word.Document.IsMasterDocument.md)|
|[IsSubdocument](./Word.Document.IsSubdocument.md)|
|[JustificationMode](./Word.Document.JustificationMode.md)|
|[KerningByAlgorithm](./Word.Document.KerningByAlgorithm.md)|
|[Kind](./Word.Document.Kind.md)|
|[LanguageDetected](./Word.Document.LanguageDetected.md)|
|[ListParagraphs](./Word.Document.ListParagraphs.md)|
|[Lists](./Word.Document.Lists.md)|
|[ListTemplates](./Word.Document.ListTemplates.md)|
|[LockQuickStyleSet](./Word.Document.LockQuickStyleSet.md)|
|[LockTheme](./Word.Document.LockTheme.md)|
|[MailEnvelope](./Word.Document.MailEnvelope.md)|
|[MailMerge](./Word.Document.MailMerge.md)|
|[Name](./Word.Document.Name.md)|
|[NoLineBreakAfter](./Word.Document.NoLineBreakAfter.md)|
|[NoLineBreakBefore](./Word.Document.NoLineBreakBefore.md)|
|[OMathBreakBin](./Word.Document.OMathBreakBin.md)|
|[OMathBreakSub](./Word.Document.OMathBreakSub.md)|
|[OMathFontName](./Word.Document.OMathFontName.md)|
|[OMathIntSubSupLim](./Word.Document.OMathIntSubSupLim.md)|
|[OMathJc](./Word.Document.OMathJc.md)|
|[OMathLeftMargin](./Word.Document.OMathLeftMargin.md)|
|[OMathNarySupSubLim](./Word.Document.OMathNarySupSubLim.md)|
|[OMathRightMargin](./Word.Document.OMathRightMargin.md)|
|[OMaths](./Word.Document.OMaths.md)|
|[OMathSmallFrac](./Word.Document.OMathSmallFrac.md)|
|[OMathWrap](./Word.Document.OMathWrap.md)|
|[OpenEncoding](./Word.Document.OpenEncoding.md)|
|[OptimizeForWord97](./Word.Document.OptimizeForWord97.md)|
|[OriginalDocumentTitle](./Word.Document.OriginalDocumentTitle.md)|
|[PageSetup](./Word.Document.PageSetup.md)|
|[Paragraphs](./Word.Document.Paragraphs.md)|
|[Parent](./Word.Document.Parent.md)|
|[Password](./Word.Document.Password.md)|
|[PasswordEncryptionAlgorithm](./Word.Document.PasswordEncryptionAlgorithm.md)|
|[PasswordEncryptionFileProperties](./Word.Document.PasswordEncryptionFileProperties.md)|
|[PasswordEncryptionKeyLength](./Word.Document.PasswordEncryptionKeyLength.md)|
|[PasswordEncryptionProvider](./Word.Document.PasswordEncryptionProvider.md)|
|[Path](./Word.Document.Path.md)|
|[Permission](./Word.Document.Permission.md)|
|[PrintFormsData](./Word.Document.PrintFormsData.md)|
|[PrintPostScriptOverText](./Word.Document.PrintPostScriptOverText.md)|
|[PrintRevisions](./Word.Document.PrintRevisions.md)|
|[ProtectionType](./Word.Document.ProtectionType.md)|
|[ReadabilityStatistics](./Word.Document.ReadabilityStatistics.md)|
|[ReadingLayoutSizeX](./Word.Document.ReadingLayoutSizeX.md)|
|[ReadingLayoutSizeY](./Word.Document.ReadingLayoutSizeY.md)|
|[ReadingModeLayoutFrozen](./Word.Document.ReadingModeLayoutFrozen.md)|
|[ReadOnly](./Word.Document.ReadOnly.md)|
|[ReadOnlyRecommended](./Word.Document.ReadOnlyRecommended.md)|
|[RemoveDateAndTime](./Word.Document.RemoveDateAndTime.md)|
|[RemovePersonalInformation](./Word.Document.RemovePersonalInformation.md)|
|[Research](./Word.Document.Research.md)|
|[RevisedDocumentTitle](./Word.Document.RevisedDocumentTitle.md)|
|[Revisions](./Word.Document.Revisions.md)|
|[Saved](./Word.Document.Saved.md)|
|[SaveEncoding](./Word.Document.SaveEncoding.md)|
|[SaveFormat](./Word.Document.SaveFormat.md)|
|[SaveFormsData](./Word.Document.SaveFormsData.md)|
|[SaveSubsetFonts](./Word.Document.SaveSubsetFonts.md)|
|[Scripts](./Word.Document.Scripts.md)|
|[Sections](./Word.Document.Sections.md)|
|[Sentences](./Word.Document.Sentences.md)|
|[ServerPolicy](./Word.Document.ServerPolicy.md)|
|[Shapes](./Word.Document.Shapes.md)|
|[ShowGrammaticalErrors](./Word.Document.ShowGrammaticalErrors.md)|
|[ShowSpellingErrors](./Word.Document.ShowSpellingErrors.md)|
|[Signatures](./Word.Document.Signatures.md)|
|[SmartDocument](./Word.Document.SmartDocument.md)|
|[SnapToGrid](./Word.Document.SnapToGrid.md)|
|[SnapToShapes](./Word.Document.SnapToShapes.md)|
|[SpellingChecked](./Word.Document.SpellingChecked.md)|
|[SpellingErrors](./Word.Document.SpellingErrors.md)|
|[StoryRanges](./Word.Document.StoryRanges.md)|
|[Styles](./Word.Document.Styles.md)|
|[StyleSheets](./Word.Document.StyleSheets.md)|
|[StyleSortMethod](./Word.Document.StyleSortMethod.md)|
|[Subdocuments](./Word.Document.Subdocuments.md)|
|[Sync](./Word.Document.Sync(property).md)|
|[Tables](./Word.Document.Tables.md)|
|[TablesOfAuthorities](./Word.Document.TablesOfAuthorities.md)|
|[TablesOfAuthoritiesCategories](./Word.Document.TablesOfAuthoritiesCategories.md)|
|[TablesOfContents](./Word.Document.TablesOfContents.md)|
|[TablesOfFigures](./Word.Document.TablesOfFigures.md)|
|[TextEncoding](./Word.Document.TextEncoding.md)|
|[TextLineEnding](./Word.Document.TextLineEnding.md)|
|[TrackFormatting](./Word.Document.TrackFormatting.md)|
|[TrackMoves](./Word.Document.TrackMoves.md)|
|[TrackRevisions](./Word.Document.TrackRevisions.md)|
|[Type](./Word.Document.Type.md)|
|[UpdateStylesOnOpen](./Word.Document.UpdateStylesOnOpen.md)|
|[UseMathDefaults](./Word.Document.UseMathDefaults.md)|
|[UserControl](./Word.Document.UserControl.md)|
|[Variables](./Word.Document.Variables.md)|
|[VBASigned](./Word.Document.VBASigned.md)|
|[VBProject](./Word.Document.VBProject.md)|
|[WebOptions](./Word.Document.WebOptions.md)|
|[Windows](./Word.Document.Windows.md)|
|[WordOpenXML](./Word.Document.WordOpenXML.md)|
|[Words](./Word.Document.Words.md)|
|[WritePassword](./Word.Document.WritePassword.md)|
|[WriteReserved](./Word.Document.WriteReserved.md)|
|[XMLSaveThroughXSLT](./Word.Document.XMLSaveThroughXSLT.md)|
|[XMLSchemaReferences](./Word.Document.XMLSchemaReferences.md)|
|[XMLShowAdvancedErrors](./Word.Document.XMLShowAdvancedErrors.md)|
|[XMLUseXSLTWhenSaving](./Word.Document.XMLUseXSLTWhenSaving.md)|

## See also


[Word Object Model Reference](./overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]