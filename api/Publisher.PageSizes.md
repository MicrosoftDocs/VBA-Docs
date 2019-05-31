---
title: PageSizes object (Publisher)
keywords: vbapb10.chm8847359
f1_keywords:
- vbapb10.chm8847359
ms.prod: publisher
api_name:
- Publisher.PageSizes
ms.assetid: f31b08cc-2c76-e2d6-d1ae-6dcf2ac5824c
ms.date: 06/01/2019
localization_priority: Normal
---


# PageSizes object (Publisher)

Represents the collection of all **[PageSize](Publisher.PageSize.md)** objects in the parent **Document** object, where each **PageSize** object represents one of the page sizes available in the current Microsoft Publisher document.


## Remarks

The page sizes represented by the **PageSizes** collection correspond to the icons displayed under **Blank Page Sizes** in the **Page Setup** dialog box in the Publisher user interface.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **PageSizes** collection to get all the page sizes available in the current document and print the list in the Immediate window.

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

- [Application](Publisher.PageSizes.Application.md)
- [Count](Publisher.PageSizes.Count.md)
- [Item](Publisher.PageSizes.Item.md)
- [Parent](Publisher.PageSizes.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]