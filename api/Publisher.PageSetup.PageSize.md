---
title: PageSetup.PageSize property (Publisher)
keywords: vbapb10.chm6946850
f1_keywords:
- vbapb10.chm6946850
ms.prod: publisher
api_name:
- Publisher.PageSetup.PageSize
ms.assetid: b0605215-5d91-e26e-d3c5-98254cf30044
ms.date: 06/12/2019
localization_priority: Normal
---


# PageSetup.PageSize property (Publisher)

Gets or sets the blank page size for the current publication. Read/write.


## Syntax

_expression_.**PageSize**

_expression_ A variable that represents a **[PageSetup](Publisher.PageSetup.md)** object.


## Return value

PageSize


## Remarks

The blank page size represented by the **PageSize** object returned or set by the **PageSize** property corresponds to one of the icons displayed under **Blank Page Sizes** in the **Page Setup** dialog box in the Microsoft Publisher user interface.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to set the blank page size for the current publication. The example sets the blank page size to Index Card, which is the blank page size at index number 130 in the **AvailablePageSizes** collection. 

See the **[AvailablePageSizes](Publisher.PageSetup.AvailablePageSizes.md)** property for an example of how to create a text file that contains the list of all page sizes available in the current publication and their corresponding index numbers.

```vb
Public Sub PageSize_Example() 
 
 ThisDocument.PageSetup.PageSize = ThisDocument.PageSetup.AvailablePageSizes.Item(130) 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]