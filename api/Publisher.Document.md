---
title: Document object (Publisher)
keywords: vbapb10.chm553713663
f1_keywords:
- vbapb10.chm553713663
ms.prod: publisher
api_name:
- Publisher.Document
ms.assetid: 44f02255-ff5b-bcfe-900f-61c8fdf61ef3
ms.date: 05/31/2019
localization_priority: Normal
---


# Document object (Publisher)

Represents a publication. 

## Remarks

Use the **[ActiveDocument](Publisher.Application.ActiveDocument.md)** property to refer to the current publication. 

## Example

This example adds a table to the first page of the active publication.

```vb
Sub NewTable() 
 With ActiveDocument.Pages(1).Shapes 
 .AddTable NumRows:=3, NumColumns:=3, Left:=72, Top:=300, _ 
 Width:=488, Height:=36 
 With .Item(1).Table.Rows(1) 
 .Cells(1).TextRange.Text = "Column1" 
 .Cells(2).TextRange.Text = "Column2" 
 .Cells(3).TextRange.Text = "Column3" 
 End With 
 End With 
End Sub
```

<br/>

You can also write the previous routine by using a reference to the **ThisDocument** module. This example uses a **ThisDocument** reference instead of **ActiveDocument**.

```vb
Sub PrintPublication() 
 With ThisDocument.Pages(1).Shapes 
 .AddTable NumRows:=3, NumColumns:=3, Left:=72, Top:=300, _ 
 Width:=488, Height:=36 
 With .Item(1).Table.Rows(1) 
 .Cells(1).TextRange.Text = "Column1" 
 .Cells(2).TextRange.Text = "Column2" 
 .Cells(3).TextRange.Text = "Column3" 
 End With 
 End With 
End Sub
```


## Events

- [BeforeClose](Publisher.Document.BeforeClose.md)
- [Open](Publisher.Document.Open.md)
- [Redo](Publisher.Document.Redo(even).md)
- [ShapesAdded](Publisher.Document.ShapesAdded.md)
- [ShapesRemoved](Publisher.Document.ShapesRemoved.md)
- [Undo](Publisher.Document.Undo(even).md)
- [WizardAfterChange](Publisher.Document.WizardAfterChange.md)

## Methods

- [BeginCustomUndoAction](Publisher.Document.BeginCustomUndoAction.md)
- [ChangeDocument](Publisher.Document.ChangeDocument.md)
- [Close](Publisher.Document.Close.md)
- [ConvertPublicationType](Publisher.Document.ConvertPublicationType.md)
- [EndCustomUndoAction](Publisher.Document.EndCustomUndoAction.md)
- [ExportAsFixedFormat](Publisher.Document.ExportAsFixedFormat.md)
- [FindShapeByWizardTag](Publisher.Document.FindShapeByWizardTag.md)
- [FindShapesByTag](Publisher.Document.FindShapesByTag.md)
- [PrintOutEx](Publisher.Document.PrintOutEx.md)
- [Redo](Publisher.Document.Redo(method).md)
- [Save](Publisher.Document.Save.md)
- [SaveAs](Publisher.Document.SaveAs.md)
- [SetBusinessInformation](Publisher.Document.SetBusinessInformation.md)
- [Undo](Publisher.Document.Undo(method).md)
- [UndoClear](Publisher.Document.UndoClear.md)
- [UpdateOLEObjects](Publisher.Document.UpdateOLEObjects.md)
- [WebPagePreview](Publisher.Document.WebPagePreview.md)

## Properties

- [ActiveView](Publisher.Document.ActiveView.md)
- [ActiveWindow](Publisher.Document.ActiveWindow.md)
- [AdvancedPrintOptions](Publisher.Document.AdvancedPrintOptions.md)
- [Application](Publisher.Document.Application.md)
- [AvailableBuildingBlocks](Publisher.document.availablebuildingblocks.md)
- [BorderArts](Publisher.Document.BorderArts.md)
- [ColorScheme](Publisher.Document.ColorScheme.md)
- [DefaultTabStop](Publisher.Document.DefaultTabStop.md)
- [DocumentDirection](Publisher.Document.DocumentDirection.md)
- [EnvelopeVisible](Publisher.Document.EnvelopeVisible.md)
- [Find](Publisher.Document.Find.md)
- [FullName](Publisher.Document.FullName.md)
- [IsDataSourceConnected](Publisher.Document.IsDataSourceConnected.md)
- [IsWizard](Publisher.Document.IsWizard.md)
- [LayoutGuides](Publisher.Document.LayoutGuides.md)
- [MailEnvelope](Publisher.Document.MailEnvelope.md)
- [MailMerge](Publisher.Document.MailMerge.md)
- [MasterPages](Publisher.Document.MasterPages.md)
- [Name](Publisher.Document.Name.md)
- [Pages](Publisher.Document.Pages.md)
- [PageSetup](Publisher.Document.PageSetup.md)
- [Parent](Publisher.Document.Parent.md)
- [Path](Publisher.Document.Path.md)
- [PrintPageBackgrounds](Publisher.Document.PrintPageBackgrounds.md)
- [PrintStyle](Publisher.Document.PrintStyle.md)
- [PublicationType](Publisher.Document.PublicationType.md)
- [ReadOnly](Publisher.Document.ReadOnly.md)
- [RedoActionsAvailable](Publisher.Document.RedoActionsAvailable.md)
- [RemovePersonalInformation](Publisher.Document.RemovePersonalInformation.md)
- [Saved](Publisher.Document.Saved.md)
- [SaveFormat](Publisher.Document.SaveFormat.md)
- [ScratchArea](Publisher.Document.ScratchArea.md)
- [Sections](Publisher.Document.Sections.md)
- [Selection](Publisher.Document.Selection.md)
- [Stories](Publisher.Document.Stories.md)
- [SurplusShapes](Publisher.Document.SurplusShapes.md)
- [Tags](Publisher.Document.Tags.md)
- [TextStyles](Publisher.Document.TextStyles.md)
- [UndoActionsAvailable](Publisher.Document.UndoActionsAvailable.md)
- [ViewBoundaries](Publisher.Document.ViewBoundaries.md)
- [ViewGuides](Publisher.Document.ViewGuides.md)
- [ViewHorizontalBaseLineGuides](Publisher.Document.ViewHorizontalBaseLineGuides.md)
- [ViewTwoPageSpread](Publisher.Document.ViewTwoPageSpread.md)
- [ViewVerticalBaseLineGuides](Publisher.Document.ViewVerticalBaseLineGuides.md)
- [WebNavigationBarSets](Publisher.Document.WebNavigationBarSets.md)
- [Wizard](Publisher.Document.Wizard.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]