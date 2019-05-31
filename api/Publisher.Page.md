---
title: Page object (Publisher)
keywords: vbapb10.chm458751
f1_keywords:
- vbapb10.chm458751
ms.prod: publisher
api_name:
- Publisher.Page
ms.assetid: 9b2e8f29-26c3-1008-0ffd-eea2147abca4
ms.date: 06/01/2019
localization_priority: Normal
---


# Page object (Publisher)

Represents a page in a publication. The **[Pages](Publisher.Pages.md)** collection contains all the **Page** objects in a publication.

## Remarks

Use **Pages** (_index_) to return a single **Page** object. 

Use the **[FindByPageID](Publisher.Pages.FindByPageID.md)** property of the **Pages** object to locate a **Page** object by using the application assigned page ID. 

Use the **[Add](Publisher.Pages.Add.md)** method to create a new page and add it to the publication. 

## Example

The following example adds new text to the first shape on the first page in the active publication.

```vb
Sub AddPageNumberField() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange 
 .InsertAfter " This text is added after the existing text." 
 .Font.Size = 15 
 End With 
End Sub
```

<br/>

The following example adds a new page to the active publication and then looks for that page by using the page ID.

```vb
Sub FindPage() 
 Dim lngPageID As Long 
 
 'Get page ID 
 lngPageID = ActiveDocument.Pages.Add(Count:=1, After:=1).PageID 
 
 'Use page ID to add a new shape to the page 
 ActiveDocument.Pages.FindByPageID(PageID:=lngPageID) _ 
 .Shapes.AddShape Type:=msoShape5pointStar, _ 
 Left:=200, Top:=72, Width:=50, Height:=50 
 
End Sub
```


## Methods

- [Delete](Publisher.Page.Delete.md)
- [Duplicate](Publisher.Page.Duplicate.md)
- [ExportEmailHTML](Publisher.Page.ExportEmailHTML.md)
- [Move](Publisher.Page.Move.md)
- [SaveAsPicture](Publisher.Page.SaveAsPicture.md)

## Properties

- [Application](Publisher.Page.Application.md)
- [Background](Publisher.Page.Background.md)
- [Footer](Publisher.Page.Footer.md)
- [Header](Publisher.Page.Header.md)
- [Height](Publisher.Page.Height.md)
- [IgnoreMaster](Publisher.Page.IgnoreMaster.md)
- [IsLeading](Publisher.Page.IsLeading.md)
- [IsTrailing](Publisher.Page.IsTrailing.md)
- [IsTwoPageMaster](Publisher.Page.IsTwoPageMaster.md)
- [IsWizardPage](Publisher.Page.IsWizardPage.md)
- [LayoutGuides](Publisher.Page.LayoutGuides.md)
- [Master](Publisher.Page.Master.md)
- [Name](Publisher.Page.Name.md)
- [PageID](Publisher.Page.PageID.md)
- [PageIndex](Publisher.Page.PageIndex.md)
- [PageNumber](Publisher.Page.PageNumber.md)
- [PageType](Publisher.Page.PageType.md)
- [Parent](Publisher.Page.Parent.md)
- [ReaderSpread](Publisher.Page.ReaderSpread.md)
- [RulerGuides](Publisher.Page.RulerGuides.md)
- [Shapes](Publisher.Page.Shapes.md)
- [Tags](Publisher.Page.Tags.md)
- [WebPageOptions](Publisher.Page.WebPageOptions.md)
- [Width](Publisher.Page.Width.md)
- [Wizard](Publisher.Page.Wizard.md)
- [XOffsetWithinReaderSpread](Publisher.Page.XOffsetWithinReaderSpread.md)
- [YOffsetWithinReaderSpread](Publisher.Page.YOffsetWithinReaderSpread.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]