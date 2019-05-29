---
title: PageSetup.RightMargin property (Word)
keywords: vbawd10.chm158400615
f1_keywords:
- vbawd10.chm158400615
ms.prod: word
api_name:
- Word.PageSetup.RightMargin
ms.assetid: abaabc8b-bb3f-fe68-ca35-d06f74d06791
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.RightMargin property (Word)

Returns or sets the distance (in points) between the right edge of the page and the right boundary of the body text. Read/write  **Single**.


## Syntax

_expression_.**RightMargin**

_expression_ A variable that represents a **[PageSetup](Word.PageSetup.md)** object.


## Remarks

If the  **[MirrorMargins](Word.PageSetup.MirrorMargins.md)** property is set to **True**, the **RightMargin** property controls the setting for outside margins and the **[LeftMargin](Word.PageSetup.LeftMargin.md)** property controls the setting for inside margins.


## Example

This example displays the right margin setting for the active document. The  **[PointsToInches](Word.Global.PointsToInches.md)** method is used to convert the result to inches.


```vb
With ActiveDocument.PageSetup 
 Msgbox "The right margin is set to " _ 
 & PointsToInches(.RightMargin) & " inches." 
End With
```

This example sets the right margin for section two in the selection. The  **[InchesToPoints](Word.Application.InchesToPoints.md)** method is used to convert inches to points.




```vb
Selection.Sections(2).PageSetup.RightMargin = InchesToPoints(1)
```


## See also


[PageSetup Object](Word.PageSetup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]