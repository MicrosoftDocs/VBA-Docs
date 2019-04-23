---
title: Options object (Word)
keywords: vbawd10.chm2487
f1_keywords:
- vbawd10.chm2487
ms.prod: word
api_name:
- Word.Options
ms.assetid: 873b7b99-3fe1-fd89-9ece-a9355cb827dc
ms.date: 06/08/2017
localization_priority: Normal
---


# Options object (Word)

Represents application and document options in Word. Many of the properties for the  **Options** object correspond to items in the **Options** dialog box.


## Remarks

Use the  **Options** property to return the **Options** object. The following example sets three application options for Word.


```vb
With Options 
 .AllowDragAndDrop = True 
 .ConfirmConversions = False 
 .MeasurementUnit = wdPoints 
End With
```


## Properties



|Name|
|:-----|
|[AddBiDirectionalMarksWhenSavingTextFile](Word.Options.AddBiDirectionalMarksWhenSavingTextFile.md)|
|[AddControlCharacters](Word.Options.AddControlCharacters.md)|
|[AddHebDoubleQuote](Word.Options.AddHebDoubleQuote.md)|
|[AlertIfNotDefault](Word.options.alertifnotdefault.md)|
|[AllowAccentedUppercase](Word.Options.AllowAccentedUppercase.md)|
|[AllowClickAndTypeMouse](Word.Options.AllowClickAndTypeMouse.md)|
|[AllowCombinedAuxiliaryForms](Word.Options.AllowCombinedAuxiliaryForms.md)|
|[AllowCompoundNounProcessing](Word.Options.AllowCompoundNounProcessing.md)|
|[AllowDragAndDrop](Word.Options.AllowDragAndDrop.md)|
|[AllowOpenInDraftView](Word.Options.AllowOpenInDraftView.md)|
|[AllowPixelUnits](Word.Options.AllowPixelUnits.md)|
|[AllowReadingMode](Word.Options.AllowReadingMode.md)|
|[AnimateScreenMovements](Word.Options.AnimateScreenMovements.md)|
|[Application](Word.Options.Application.md)|
|[ApplyFarEastFontsToAscii](Word.Options.ApplyFarEastFontsToAscii.md)|
|[ArabicMode](Word.Options.ArabicMode.md)|
|[ArabicNumeral](Word.Options.ArabicNumeral.md)|
|[AutoCreateNewDrawings](Word.Options.AutoCreateNewDrawings.md)|
|[AutoFormatApplyBulletedLists](Word.Options.AutoFormatApplyBulletedLists.md)|
|[AutoFormatApplyFirstIndents](Word.Options.AutoFormatApplyFirstIndents.md)|
|[AutoFormatApplyHeadings](Word.Options.AutoFormatApplyHeadings.md)|
|[AutoFormatApplyLists](Word.Options.AutoFormatApplyLists.md)|
|[AutoFormatApplyOtherParas](Word.Options.AutoFormatApplyOtherParas.md)|
|[AutoFormatAsYouTypeApplyBorders](Word.Options.AutoFormatAsYouTypeApplyBorders.md)|
|[AutoFormatAsYouTypeApplyBulletedLists](Word.Options.AutoFormatAsYouTypeApplyBulletedLists.md)|
|[AutoFormatAsYouTypeApplyClosings](Word.Options.AutoFormatAsYouTypeApplyClosings.md)|
|[AutoFormatAsYouTypeApplyDates](Word.Options.AutoFormatAsYouTypeApplyDates.md)|
|[AutoFormatAsYouTypeApplyFirstIndents](Word.Options.AutoFormatAsYouTypeApplyFirstIndents.md)|
|[AutoFormatAsYouTypeApplyHeadings](Word.Options.AutoFormatAsYouTypeApplyHeadings.md)|
|[AutoFormatAsYouTypeApplyNumberedLists](Word.Options.AutoFormatAsYouTypeApplyNumberedLists.md)|
|[AutoFormatAsYouTypeApplyTables](Word.Options.AutoFormatAsYouTypeApplyTables.md)|
|[AutoFormatAsYouTypeAutoLetterWizard](Word.Options.AutoFormatAsYouTypeAutoLetterWizard.md)|
|[AutoFormatAsYouTypeDefineStyles](Word.Options.AutoFormatAsYouTypeDefineStyles.md)|
|[AutoFormatAsYouTypeDeleteAutoSpaces](Word.Options.AutoFormatAsYouTypeDeleteAutoSpaces.md)|
|[AutoFormatAsYouTypeFormatListItemBeginning](Word.Options.AutoFormatAsYouTypeFormatListItemBeginning.md)|
|[AutoFormatAsYouTypeInsertClosings](Word.Options.AutoFormatAsYouTypeInsertClosings.md)|
|[AutoFormatAsYouTypeInsertOvers](Word.Options.AutoFormatAsYouTypeInsertOvers.md)|
|[AutoFormatAsYouTypeMatchParentheses](Word.Options.AutoFormatAsYouTypeMatchParentheses.md)|
|[AutoFormatAsYouTypeReplaceFarEastDashes](Word.Options.AutoFormatAsYouTypeReplaceFarEastDashes.md)|
|[AutoFormatAsYouTypeReplaceFractions](Word.Options.AutoFormatAsYouTypeReplaceFractions.md)|
|[AutoFormatAsYouTypeReplaceHyperlinks](Word.Options.AutoFormatAsYouTypeReplaceHyperlinks.md)|
|[AutoFormatAsYouTypeReplaceOrdinals](Word.Options.AutoFormatAsYouTypeReplaceOrdinals.md)|
|[AutoFormatAsYouTypeReplacePlainTextEmphasis](Word.Options.AutoFormatAsYouTypeReplacePlainTextEmphasis.md)|
|[AutoFormatAsYouTypeReplaceQuotes](Word.Options.AutoFormatAsYouTypeReplaceQuotes.md)|
|[AutoFormatAsYouTypeReplaceSymbols](Word.Options.AutoFormatAsYouTypeReplaceSymbols.md)|
|[AutoFormatDeleteAutoSpaces](Word.Options.AutoFormatDeleteAutoSpaces.md)|
|[AutoFormatMatchParentheses](Word.Options.AutoFormatMatchParentheses.md)|
|[AutoFormatPlainTextWordMail](Word.Options.AutoFormatPlainTextWordMail.md)|
|[AutoFormatPreserveStyles](Word.Options.AutoFormatPreserveStyles.md)|
|[AutoFormatReplaceFarEastDashes](Word.Options.AutoFormatReplaceFarEastDashes.md)|
|[AutoFormatReplaceFractions](Word.Options.AutoFormatReplaceFractions.md)|
|[AutoFormatReplaceHyperlinks](Word.Options.AutoFormatReplaceHyperlinks.md)|
|[AutoFormatReplaceOrdinals](Word.Options.AutoFormatReplaceOrdinals.md)|
|[AutoFormatReplacePlainTextEmphasis](Word.Options.AutoFormatReplacePlainTextEmphasis.md)|
|[AutoFormatReplaceQuotes](Word.Options.AutoFormatReplaceQuotes.md)|
|[AutoFormatReplaceSymbols](Word.Options.AutoFormatReplaceSymbols.md)|
|[AutoKeyboardSwitching](Word.Options.AutoKeyboardSwitching.md)|
|[AutoWordSelection](Word.Options.AutoWordSelection.md)|
|[BackgroundSave](Word.Options.BackgroundSave.md)|
|[BibliographySort](Word.Options.BibliographySort.md)|
|[BibliographyStyle](Word.Options.BibliographyStyle.md)|
|[BrazilReform](Word.Options.BrazilReform.md)|
|[ButtonFieldClicks](Word.Options.ButtonFieldClicks.md)|
|[CheckGrammarAsYouType](Word.Options.CheckGrammarAsYouType.md)|
|[CheckGrammarWithSpelling](Word.Options.CheckGrammarWithSpelling.md)|
|[CheckHangulEndings](Word.Options.CheckHangulEndings.md)|
|[CheckSpellingAsYouType](Word.Options.CheckSpellingAsYouType.md)|
|[CloudSignInOption](Word.options.cloudsigninoption.md)|
|[CommentsColor](Word.Options.CommentsColor.md)|
|[ConfirmConversions](Word.Options.ConfirmConversions.md)|
|[ContextualSpeller](Word.Options.ContextualSpeller.md)|
|[ConvertHighAnsiToFarEast](Word.Options.ConvertHighAnsiToFarEast.md)|
|[CreateBackup](Word.Options.CreateBackup.md)|
|[Creator](Word.Options.Creator.md)|
|[CtrlClickHyperlinkToOpen](Word.Options.CtrlClickHyperlinkToOpen.md)|
|[CursorMovement](Word.Options.CursorMovement.md)|
|[DefaultBorderColor](Word.Options.DefaultBorderColor.md)|
|[DefaultBorderColorIndex](Word.Options.DefaultBorderColorIndex.md)|
|[DefaultBorderLineStyle](Word.Options.DefaultBorderLineStyle.md)|
|[DefaultBorderLineWidth](Word.Options.DefaultBorderLineWidth.md)|
|[DefaultEPostageApp](Word.Options.DefaultEPostageApp.md)|
|[DefaultFilePath](Word.Options.DefaultFilePath.md)|
|[DefaultHighlightColorIndex](Word.Options.DefaultHighlightColorIndex.md)|
|[DefaultOpenFormat](Word.Options.DefaultOpenFormat.md)|
|[DefaultTextEncoding](Word.Options.DefaultTextEncoding.md)|
|[DefaultTray](Word.Options.DefaultTray.md)|
|[DefaultTrayID](Word.Options.DefaultTrayID.md)|
|[DeletedCellColor](Word.Options.DeletedCellColor.md)|
|[DeletedTextColor](Word.Options.DeletedTextColor.md)|
|[DeletedTextMark](Word.Options.DeletedTextMark.md)|
|[DiacriticColorVal](Word.Options.DiacriticColorVal.md)|
|[DisableFeaturesbyDefault](Word.Options.DisableFeaturesbyDefault.md)|
|[DisableFeaturesIntroducedAfterbyDefault](Word.Options.DisableFeaturesIntroducedAfterbyDefault.md)|
|[DisplayAlignmentGuides](Word.options.displayalignmentguides.md)|
|[DisplayGridLines](Word.Options.DisplayGridLines.md)|
|[DisplayPasteOptions](Word.Options.DisplayPasteOptions.md)|
|[DocumentViewDirection](Word.Options.DocumentViewDirection.md)|
|[DoNotPromptForConvert](Word.Options.DoNotPromptForConvert.md)|
|[EnableHangulHanjaRecentOrdering](Word.Options.EnableHangulHanjaRecentOrdering.md)|
|[EnableLegacyIMEMode](Word.Options.EnableLegacyIMEMode.md)|
|[EnableLiveDrag](Word.options.enablelivedrag.md)|
|[EnableLivePreview](Word.Options.EnableLivePreview.md)|
|[EnableMisusedWordsDictionary](Word.Options.EnableMisusedWordsDictionary.md)|
|[EnableProofingToolsAdvertisement](Word.options.enableproofingtoolsadvertisement.md)|
|[EnableSound](Word.Options.EnableSound.md)|
|[EnvelopeFeederInstalled](Word.Options.EnvelopeFeederInstalled.md)|
|[ExpandHeadingsOnOpen](Word.options.expandheadingsonopen.md)|
|[FormatScanning](Word.Options.FormatScanning.md)|
|[FrenchReform](Word.Options.FrenchReform.md)|
|[GridDistanceHorizontal](Word.Options.GridDistanceHorizontal.md)|
|[GridDistanceVertical](Word.Options.GridDistanceVertical.md)|
|[GridOriginHorizontal](Word.Options.GridOriginHorizontal.md)|
|[GridOriginVertical](Word.Options.GridOriginVertical.md)|
|[HangulHanjaFastConversion](Word.Options.HangulHanjaFastConversion.md)|
|[HebrewMode](Word.Options.HebrewMode.md)|
|[IgnoreInternetAndFileAddresses](Word.Options.IgnoreInternetAndFileAddresses.md)|
|[IgnoreMixedDigits](Word.Options.IgnoreMixedDigits.md)|
|[IgnoreUppercase](Word.Options.IgnoreUppercase.md)|
|[IMEAutomaticControl](Word.Options.IMEAutomaticControl.md)|
|[InlineConversion](Word.Options.InlineConversion.md)|
|[InsertedCellColor](Word.Options.InsertedCellColor.md)|
|[InsertedTextColor](Word.Options.InsertedTextColor.md)|
|[InsertedTextMark](Word.Options.InsertedTextMark.md)|
|[INSKeyForOvertype](Word.Options.INSKeyForOvertype.md)|
|[INSKeyForPaste](Word.Options.INSKeyForPaste.md)|
|[InterpretHighAnsi](Word.Options.InterpretHighAnsi.md)|
|[LocalNetworkFile](Word.Options.LocalNetworkFile.md)|
|[MapPaperSize](Word.Options.MapPaperSize.md)|
|[MarginAlignmentGuides](Word.options.marginalignmentguides.md)|
|[MatchFuzzyAY](Word.Options.MatchFuzzyAY.md)|
|[MatchFuzzyBV](Word.Options.MatchFuzzyBV.md)|
|[MatchFuzzyByte](Word.Options.MatchFuzzyByte.md)|
|[MatchFuzzyCase](Word.Options.MatchFuzzyCase.md)|
|[MatchFuzzyDash](Word.Options.MatchFuzzyDash.md)|
|[MatchFuzzyDZ](Word.Options.MatchFuzzyDZ.md)|
|[MatchFuzzyHF](Word.Options.MatchFuzzyHF.md)|
|[MatchFuzzyHiragana](Word.Options.MatchFuzzyHiragana.md)|
|[MatchFuzzyIterationMark](Word.Options.MatchFuzzyIterationMark.md)|
|[MatchFuzzyKanji](Word.Options.MatchFuzzyKanji.md)|
|[MatchFuzzyKiKu](Word.Options.MatchFuzzyKiKu.md)|
|[MatchFuzzyOldKana](Word.Options.MatchFuzzyOldKana.md)|
|[MatchFuzzyProlongedSoundMark](Word.Options.MatchFuzzyProlongedSoundMark.md)|
|[MatchFuzzyPunctuation](Word.Options.MatchFuzzyPunctuation.md)|
|[MatchFuzzySmallKana](Word.Options.MatchFuzzySmallKana.md)|
|[MatchFuzzySpace](Word.Options.MatchFuzzySpace.md)|
|[MatchFuzzyTC](Word.Options.MatchFuzzyTC.md)|
|[MatchFuzzyZJ](Word.Options.MatchFuzzyZJ.md)|
|[MeasurementUnit](Word.Options.MeasurementUnit.md)|
|[MergedCellColor](Word.Options.MergedCellColor.md)|
|[MonthNames](Word.Options.MonthNames.md)|
|[MoveFromTextColor](Word.Options.MoveFromTextColor.md)|
|[MoveFromTextMark](Word.Options.MoveFromTextMark.md)|
|[MoveToTextColor](Word.Options.MoveToTextColor.md)|
|[MoveToTextMark](Word.Options.MoveToTextMark.md)|
|[MultipleWordConversionsMode](Word.Options.MultipleWordConversionsMode.md)|
|[OMathAutoBuildUp](Word.Options.OMathAutoBuildUp.md)|
|[OMathCopyLF](Word.Options.OMathCopyLF.md)|
|[OptimizeForWord97byDefault](Word.Options.OptimizeForWord97byDefault.md)|
|[Overtype](Word.Options.Overtype.md)|
|[PageAlignmentGuides](Word.options.pagealignmentguides.md)|
|[Pagination](Word.Options.Pagination.md)|
|[ParagraphAlignmentGuides](Word.options.paragraphalignmentguides.md)|
|[Parent](Word.Options.Parent.md)|
|[PasteAdjustParagraphSpacing](Word.Options.PasteAdjustParagraphSpacing.md)|
|[PasteAdjustTableFormatting](Word.Options.PasteAdjustTableFormatting.md)|
|[PasteAdjustWordSpacing](Word.Options.PasteAdjustWordSpacing.md)|
|[PasteFormatBetweenDocuments](Word.Options.PasteFormatBetweenDocuments.md)|
|[PasteFormatBetweenStyledDocuments](Word.Options.PasteFormatBetweenStyledDocuments.md)|
|[PasteFormatFromExternalSource](Word.Options.PasteFormatFromExternalSource.md)|
|[PasteFormatWithinDocument](Word.Options.PasteFormatWithinDocument.md)|
|[PasteMergeFromPPT](Word.Options.PasteMergeFromPPT.md)|
|[PasteMergeFromXL](Word.Options.PasteMergeFromXL.md)|
|[PasteMergeLists](Word.Options.PasteMergeLists.md)|
|[PasteOptionKeepBulletsAndNumbers](Word.Options.PasteOptionKeepBulletsAndNumbers.md)|
|[PasteSmartCutPaste](Word.Options.PasteSmartCutPaste.md)|
|[PasteSmartStyleBehavior](Word.Options.PasteSmartStyleBehavior.md)|
|[PictureEditor](Word.Options.PictureEditor.md)|
|[PictureWrapType](Word.Options.PictureWrapType.md)|
|[PortugalReform](Word.Options.PortugalReform.md)|
|[PrecisePositioning](Word.Options.PrecisePositioning.md)|
|[PreferCloudSaveLocations](Word.options.prefercloudsavelocations.md)|
|[PrintBackground](Word.Options.PrintBackground.md)|
|[PrintBackgrounds](Word.Options.PrintBackgrounds.md)|
|[PrintComments](Word.Options.PrintComments.md)|
|[PrintDraft](Word.Options.PrintDraft.md)|
|[PrintDrawingObjects](Word.Options.PrintDrawingObjects.md)|
|[PrintEvenPagesInAscendingOrder](Word.Options.PrintEvenPagesInAscendingOrder.md)|
|[PrintFieldCodes](Word.Options.PrintFieldCodes.md)|
|[PrintHiddenText](Word.Options.PrintHiddenText.md)|
|[PrintOddPagesInAscendingOrder](Word.Options.PrintOddPagesInAscendingOrder.md)|
|[PrintProperties](Word.Options.PrintProperties.md)|
|[PrintReverse](Word.Options.PrintReverse.md)|
|[PrintXMLTag](Word.Options.PrintXMLTag.md)|
|[PromptUpdateStyle](Word.Options.PromptUpdateStyle.md)|
|[RepeatWord](Word.Options.RepeatWord.md)|
|[ReplaceSelection](Word.Options.ReplaceSelection.md)|
|[RevisedLinesColor](Word.Options.RevisedLinesColor.md)|
|[RevisedLinesMark](Word.Options.RevisedLinesMark.md)|
|[RevisedPropertiesColor](Word.Options.RevisedPropertiesColor.md)|
|[RevisedPropertiesMark](Word.Options.RevisedPropertiesMark.md)|
|[RevisionsBalloonPrintOrientation](Word.Options.RevisionsBalloonPrintOrientation.md)|
|[RTFInClipboard](Word.options.rtfinclipboard.md)|
|[SaveInterval](Word.Options.SaveInterval.md)|
|[SaveNormalPrompt](Word.Options.SaveNormalPrompt.md)|
|[SavePropertiesPrompt](Word.Options.SavePropertiesPrompt.md)|
|[SendMailAttach](Word.Options.SendMailAttach.md)|
|[SequenceCheck](Word.Options.SequenceCheck.md)|
|[ShortMenuNames](Word.Options.Options.ShortMenuNames.md)|
|[ShowControlCharacters](Word.Options.ShowControlCharacters.md)|
|[ShowDevTools](Word.Options.ShowDevTools.md)|
|[ShowDiacritics](Word.Options.ShowDiacritics.md)|
|[ShowFormatError](Word.Options.ShowFormatError.md)|
|[ShowMarkupOpenSave](Word.Options.ShowMarkupOpenSave.md)|
|[ShowMenuFloaties](Word.Options.ShowMenuFloaties.md)|
|[ShowReadabilityStatistics](Word.Options.ShowReadabilityStatistics.md)|
|[ShowSelectionFloaties](Word.Options.ShowSelectionFloaties.md)|
|[SmartCursoring](Word.Options.SmartCursoring.md)|
|[SmartCutPaste](Word.Options.SmartCutPaste.md)|
|[SmartParaSelection](Word.Options.SmartParaSelection.md)|
|[SnapToGrid](Word.Options.SnapToGrid.md)|
|[SnapToShapes](Word.Options.SnapToShapes.md)|
|[SpanishMode](Word.Options.SpanishMode.md)|
|[SplitCellColor](Word.Options.SplitCellColor.md)|
|[StoreRSIDOnSave](Word.Options.StoreRSIDOnSave.md)|
|[StrictFinalYaa](Word.Options.StrictFinalYaa.md)|
|[StrictInitialAlefHamza](Word.Options.StrictInitialAlefHamza.md)|
|[StrictRussianE](Word.Options.StrictRussianE.md)|
|[StrictTaaMarboota](Word.Options.StrictTaaMarboota.md)|
|[SuggestFromMainDictionaryOnly](Word.Options.SuggestFromMainDictionaryOnly.md)|
|[SuggestSpellingCorrections](Word.Options.SuggestSpellingCorrections.md)|
|[TabIndentKey](Word.Options.TabIndentKey.md)|
|[TypeNReplace](Word.Options.TypeNReplace.md)|
|[UpdateFieldsAtPrint](Word.Options.UpdateFieldsAtPrint.md)|
|[UpdateFieldsWithTrackedChangesAtPrint](Word.Options.UpdateFieldsWithTrackedChangesAtPrint.md)|
|[UpdateLinksAtOpen](Word.Options.UpdateLinksAtOpen.md)|
|[UpdateLinksAtPrint](Word.Options.UpdateLinksAtPrint.md)|
|[UpdateStyleListBehavior](Word.Options.UpdateStyleListBehavior.md)|
|[UseCharacterUnit](Word.Options.UseCharacterUnit.md)|
|[UseDiffDiacColor](Word.Options.UseDiffDiacColor.md)|
|[UseGermanSpellingReform](Word.Options.UseGermanSpellingReform.md)|
|[UseLocalUserInfo](Word.options.uselocaluserinfo.md)|
|[UseNormalStyleForList](Word.Options.UseNormalStyleForList.md)|
|[UseSubPixelPositioning](Word.options.usesubpixelpositioning.md)|
|[VisualSelection](Word.Options.VisualSelection.md)|
|[WarnBeforeSavingPrintingSendingMarkup](Word.Options.WarnBeforeSavingPrintingSendingMarkup.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]