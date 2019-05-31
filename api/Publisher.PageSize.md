---
title: PageSize object (Publisher)
keywords: vbapb10.chm8912895
f1_keywords:
- vbapb10.chm8912895
ms.prod: publisher
api_name:
- Publisher.PageSize
ms.assetid: 80767524-6f0c-0d3f-388a-a38891b2d04a
ms.date: 06/01/2019
localization_priority: Normal
---


# PageSize object (Publisher)

Represents the page size of the current Microsoft Publisher publication.


## Remarks

The page size represented by the **PageSize** object corresponds to one of the icons displayed under **Blank Page Sizes** in the **Page Setup** dialog box in the Publisher user interface.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **Name** property of the **PageSize** object to get a list of the names of all the page sizes available in the current document and print the list in the Immediate window.

```vb
Public Sub PageSizes_Example() 
 
 Dim pubPageSizes As Publisher.PageSizes 
 Dim pubPageSize As Publisher.PageSize 
 
 Set pubPageSizes = ThisDocument.PageSetup.AvailablePageSizes 
 For Each pubPageSize In pubPageSizes 
 Debug.Print pubPageSize.Name 
 Next 
 
End Sub
```


## Properties

- [Application](Publisher.PageSize.Application.md)
- [HasBackgroundImage](Publisher.PageSize.HasBackgroundImage.md)
- [HorizontalGap](Publisher.PageSize.HorizontalGap.md)
- [LeftMargin](Publisher.PageSize.LeftMargin.md)
- [Name](Publisher.PageSize.Name.md)
- [PageHeight](Publisher.PageSize.PageHeight.md)
- [PageWidth](Publisher.PageSize.PageWidth.md)
- [Parent](Publisher.PageSize.Parent.md)
- [TopMargin](Publisher.PageSize.TopMargin.md)
- [VerticalGap](Publisher.PageSize.VerticalGap.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]