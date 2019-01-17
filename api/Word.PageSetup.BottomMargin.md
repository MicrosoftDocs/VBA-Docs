---
title: PageSetup.BottomMargin property (Word)
keywords: vbawd10.chm158400613
f1_keywords:
- vbawd10.chm158400613
ms.prod: word
api_name:
- Word.PageSetup.BottomMargin
ms.assetid: 2633c609-3f16-583b-ba80-dddf4dcd8b71
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.BottomMargin property (Word)

Returns or sets the distance (in points) between the bottom edge of the page and the bottom boundary of the body text. Read/write  **Single**.


## Syntax

 _expression_. `BottomMargin`

 _expression_ A variable that represents a '[PageSetup](Word.PageSetup.md)' object.


## Example

This example sets the bottom margin to 72 points (1 inch) and the top margin to 2 inches for the active document. The  **[InchesToPoints](Word.Application.InchesToPoints.md)** method is used to convert inches to points.


```vb
With ActiveDocument.PageSetup 
 .BottomMargin = 72 
 .TopMargin = InchesToPoints(2) 
End With
```

This example sets the bottom margin to 2.5 inches for all the sections in the current selection.




```vb
Selection.PageSetup.BottomMargin = InchesToPoints(2.5)
```

This example returns the bottom margin for section 1 in the selection. The  **[PointsToInches](Word.Global.PointsToInches.md)** method is used to convert the result to inches.




```vb
Dim sngMargin As Single 
 
sngMargin = Selection.Sections(1).PageSetup.BottomMargin 
MsgBox PointsToInches(sngMargin) & " inches"
```


## See also


[PageSetup Object](Word.PageSetup.md)

