---
title: Global object (Word)
keywords: vbawd10.chm2489
f1_keywords:
- vbawd10.chm2489
ms.prod: word
api_name:
- Word.Global
ms.assetid: b91e7459-08d5-ea8c-42e0-f7b9bfd1a72c
ms.date: 06/08/2017
localization_priority: Normal
---


# Global object (Word)

Contains top-level properties and methods that do not need to be preceded by the  **Application** property.


## Remarks

The following two statements have the same result. One statement uses the  **Application** property to access the **Documents** collection, and one does not. Both statements are equal and achieve the same result.


```vb
Documents(1).Content.Bold = True 
Application.Documents(1).Content.Bold = True
```


## Methods



|Name|
|:-----|
|[BuildKeyCode](Word.Global.BuildKeyCode.md)|
|[CentimetersToPoints](Word.Global.CentimetersToPoints.md)|
|[ChangeFileOpenDirectory](Word.Global.ChangeFileOpenDirectory.md)|
|[CheckSpelling](Word.Global.CheckSpelling.md)|
|[CleanString](Word.Global.CleanString.md)|
|[DDEExecute](Word.Global.DDEExecute.md)|
|[DDEInitiate](Word.Global.DDEInitiate.md)|
|[DDEPoke](Word.Global.DDEPoke.md)|
|[DDERequest](Word.Global.DDERequest.md)|
|[DDETerminate](Word.Global.DDETerminate.md)|
|[DDETerminateAll](Word.Global.DDETerminateAll.md)|
|[GetSpellingSuggestions](Word.Global.GetSpellingSuggestions.md)|
|[Help](Word.Global.Help.md)|
|[InchesToPoints](Word.Global.InchesToPoints.md)|
|[KeyString](Word.Global.KeyString.md)|
|[LinesToPoints](Word.Global.LinesToPoints.md)|
|[MillimetersToPoints](Word.Global.MillimetersToPoints.md)|
|[NewWindow](Word.Global.NewWindow.md)|
|[PicasToPoints](Word.Global.PicasToPoints.md)|
|[PixelsToPoints](Word.Global.PixelsToPoints.md)|
|[PointsToCentimeters](Word.Global.PointsToCentimeters.md)|
|[PointsToInches](Word.Global.PointsToInches.md)|
|[PointsToLines](Word.Global.PointsToLines.md)|
|[PointsToMillimeters](Word.Global.PointsToMillimeters.md)|
|[PointsToPicas](Word.Global.PointsToPicas.md)|
|[PointsToPixels](Word.Global.PointsToPixels.md)|
|[Repeat](Word.Global.Repeat.md)|

## Properties



|Name|
|:-----|
|[ActiveDocument](Word.Global.ActiveDocument.md)|
|[ActivePrinter](Word.Global.ActivePrinter.md)|
|[ActiveProtectedViewWindow](Word.Global.ActiveProtectedViewWindow.md)|
|[ActiveWindow](Word.Global.ActiveWindow.md)|
|[AddIns](Word.Global.AddIns.md)|
|[Application](Word.Global.Application.md)|
|[AutoCaptions](Word.Global.AutoCaptions.md)|
|[AutoCorrect](Word.Global.AutoCorrect.md)|
|[AutoCorrectEmail](Word.Global.AutoCorrectEmail.md)|
|[CaptionLabels](Word.Global.CaptionLabels.md)|
|[CommandBars](Word.Global.CommandBars.md)|
|[Creator](Word.Global.Creator.md)|
|[CustomDictionaries](Word.Global.CustomDictionaries.md)|
|[CustomizationContext](Word.Global.CustomizationContext.md)|
|[Dialogs](Word.Global.Dialogs.md)|
|[Documents](Word.Global.Documents.md)|
|[FileConverters](Word.Global.FileConverters.md)|
|[FindKey](Word.Global.FindKey.md)|
|[FontNames](Word.Global.FontNames.md)|
|[HangulHanjaDictionaries](Word.Global.HangulHanjaDictionaries.md)|
|[IsObjectValid](Word.Global.IsObjectValid.md)|
|[IsSandboxed](Word.Global.IsSandboxed.md)|
|[KeyBindings](Word.Global.KeyBindings.md)|
|[KeysBoundTo](Word.Global.KeysBoundTo.md)|
|[LandscapeFontNames](Word.Global.LandscapeFontNames.md)|
|[Languages](Word.Global.Languages.md)|
|[LanguageSettings](Word.Global.LanguageSettings.md)|
|[ListGalleries](Word.Global.ListGalleries.md)|
|[MacroContainer](Word.Global.MacroContainer.md)|
|[Name](Word.Global.Name.md)|
|[NormalTemplate](Word.Global.NormalTemplate.md)|
|[Options](Word.Global.Options.md)|
|[Parent](Word.Global.Parent.md)|
|[PortraitFontNames](Word.Global.PortraitFontNames.md)|
|[PrintPreview](Word.Global.PrintPreview.md)|
|[ProtectedViewWindows](Word.Global.ProtectedViewWindows.md)|
|[RecentFiles](Word.Global.RecentFiles.md)|
|[Selection](Word.Global.Selection.md)|
|[ShowVisualBasicEditor](Word.Global.ShowVisualBasicEditor.md)|
|[StatusBar](Word.Global.StatusBar.md)|
|[SynonymInfo](Word.Global.SynonymInfo.md)|
|[System](Word.Global.System.md)|
|[Tasks](Word.Global.Tasks.md)|
|[Templates](Word.Global.Templates.md)|
|[VBE](Word.Global.VBE.md)|
|[Windows](Word.Global.Windows.md)|
|[WordBasic](Word.Global.WordBasic.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]