---
title: HTMLDivisions object (Word)
keywords: vbawd10.chm2536
f1_keywords:
- vbawd10.chm2536
ms.prod: word
api_name:
- Word.HTMLDivisions
ms.assetid: fe896440-817f-5485-794c-c5e9700cd062
ms.date: 06/08/2017
localization_priority: Normal
---


# HTMLDivisions object (Word)

A collection of  **HTMLDivision** objects that represents the HTML DIV elements that exist in a web document.


## Remarks

Use the **HTMLDivisions** property to return the **HTMLDivisions** collection. Use the **Add** method to add an HTML division to a web document. This example adds a new HTML division to the active document, adds text to the division, and formats the borders around the division.


```vb
Sub NewDivision() 
 
 With ActiveDocument.HTMLDivisions 
 .Add 
 .Item(Index:=1).Range.Text = "This is a new HTML division." 
 With .Item(1) 
 With .Borders(wdBorderBottom) 
 .LineStyle = wdLineStyleTriple 
 .LineWidth = wdLineWidth025pt 
 .Color = wdColorRed 
 End With 
 With .Borders(wdBorderTop) 
 .LineStyle = wdLineStyleDot 
 .LineWidth = wdLineWidth050pt 
 .Color = wdColorBlue 
 End With 
 With .Borders(wdBorderLeft) 
 .LineStyle = wdLineStyleDouble 
 .LineWidth = wdLineWidth075pt 
 .Color = wdColorBrightGreen 
 End With 
 With .Borders(wdBorderRight) 
 .LineStyle = wdLineStyleDashDotDot 
 .LineWidth = wdLineWidth075pt 
 .Color = wdColorTurquoise 
 End With 
 End With 
 End With 
 
End Sub
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]