---
title: HTMLDivision object (Word)
keywords: vbawd10.chm2535
f1_keywords:
- vbawd10.chm2535
ms.prod: word
api_name:
- Word.HTMLDivision
ms.assetid: a38918ed-61aa-3fd1-3522-d077f1ff312f
ms.date: 06/08/2017
localization_priority: Normal
---


# HTMLDivision object (Word)

Represents a single HTML DIV element within a web document. The  **HTMLDivision** object is a member of the **HTMLDivisions** collection.


## Remarks

Use  **HTMLDivisions** (Index), where Index refers to the HTML division in the document, to return a single **HTMLDivision** object. Use the **Borders** property to format border properties for an HTML division. This example formats three nested divisions in the active document. This example assumes that the active document is an HTML document with at least three divisions.


```vb
Sub FormatHTMLDivisions() 
 With ActiveDocument.HTMLDivisions(1) 
 With .Borders(wdBorderLeft) 
 .Color = wdColorRed 
 .LineStyle = wdLineStyleSingle 
 End With 
 With .Borders(wdBorderTop) 
 .Color = wdColorRed 
 .LineStyle = wdLineStyleSingle 
 End With 
 With .HTMLDivisions(1) 
 .LeftIndent = InchesToPoints(1) 
 .RightIndent = InchesToPoints(1) 
 With .Borders(wdBorderRight) 
 .Color = wdColorBlue 
 .LineStyle = wdLineStyleDouble 
 End With 
 End With 
 With .Borders(wdBorderBottom) 
 .Color = wdColorBlue 
 .LineStyle = wdLineStyleDouble 
 End With 
 With .HTMLDivisions(1) 
 .LeftIndent = InchesToPoints(1) 
 .RightIndent = InchesToPoints(1) 
 With .Borders(wdBorderLeft) 
 .Color = wdColorBlack 
 .LineStyle = wdLineStyleDot 
 End With 
 With .Borders(wdBorderTop) 
 .Color = wdColorBlack 
 .LineStyle = wdLineStyleDot 
 End With 
 End With 
 End With 
 End With 
 
End Sub
```

HTML divisions can be nested within multiple HTML divisions. Use the  **HTMLDivisionParent** method to access a parent HTML division of the current HTML division. This example formats the borders for two HTML divisions in the active document. This example assumes that the active document is an HTML document with at least two divisions.




```vb
Sub FormatHTMLDivisions() 
 With ActiveDocument.HTMLDivisions(1) 
 With .HTMLDivisions(1) 
 .LeftIndent = InchesToPoints(1) 
 .RightIndent = InchesToPoints(1) 
 With .Borders(wdBorderLeft) 
 .Color = wdColorBlue 
 .LineStyle = wdLineStyleDouble 
 End With 
 With .Borders(wdBorderRight) 
 .Color = wdColorBlue 
 .LineStyle = wdLineStyleDouble 
 End With 
 With .HTMLDivisionParent 
 .LeftIndent = InchesToPoints(1) 
 .RightIndent = InchesToPoints(1) 
 With .Borders(wdBorderTop) 
 .Color = wdColorBlack 
 .LineStyle = wdLineStyleDot 
 End With 
 With .Borders(wdBorderBottom) 
 .Color = wdColorBlack 
 .LineStyle = wdLineStyleDot 
 End With 
 End With 
 End With 
 End With 
End Sub
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]