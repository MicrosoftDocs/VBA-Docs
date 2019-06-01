---
title: PageSetup object (Publisher)
keywords: vbapb10.chm7012351
f1_keywords:
- vbapb10.chm7012351
ms.prod: publisher
api_name:
- Publisher.PageSetup
ms.assetid: 23fe3235-88ea-ac93-6d7d-850298263046
ms.date: 06/01/2019
localization_priority: Normal
---


# PageSetup object (Publisher)

Contains information about the page setup for the pages in a publication.

## Remarks

Use the **[Document.PageSetup](Publisher.Document.PageSetup.md)** property to return the **PageSetup** object. 

## Example

The following example sets all pages in the active publication to be 8.5 inches wide and 11 inches high.

```vb
Sub SetPageSetupOptions() 
 With ActiveDocument.PageSetup 
 .PageHeight = 11 * 72 
 .PageWidth = 8.5 * 72 
 End With 
End Sub
```


## Properties

- [Application](Publisher.PageSetup.Application.md)
- [AvailablePageSizes](Publisher.PageSetup.AvailablePageSizes.md)
- [HorizontalGap](Publisher.PageSetup.HorizontalGap.md)
- [LeftMargin](Publisher.PageSetup.LeftMargin.md)
- [PageHeight](Publisher.PageSetup.PageHeight.md)
- [PageSize](Publisher.PageSetup.PageSize.md)
- [PageWidth](Publisher.PageSetup.PageWidth.md)
- [Parent](Publisher.PageSetup.Parent.md)
- [PublicationLayout](Publisher.pagesetup.publicationlayout.md)
- [TopMargin](Publisher.PageSetup.TopMargin.md)
- [VerticalGap](Publisher.PageSetup.VerticalGap.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]