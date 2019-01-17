---
title: LineFormat.EndArrowheadLength property (Word)
keywords: vbawd10.chm164233321
f1_keywords:
- vbawd10.chm164233321
ms.prod: word
api_name:
- Word.LineFormat.EndArrowheadLength
ms.assetid: 70aa1917-01ed-8a1c-a910-bb7f1175fd52
ms.date: 06/08/2017
localization_priority: Normal
---


# LineFormat.EndArrowheadLength property (Word)

Returns or sets the length of the arrowhead at the end of the specified line. Read/write  **MsoArrowheadLength**.


## Syntax

 _expression_. `EndArrowheadLength`

 _expression_ Required. A variable that represents a '[LineFormat](Word.LineFormat.md)' object.


## Example

This example adds a line to the active document. There is a short, narrow oval on the line's starting point and a long, wide triangle on its endpoint.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 
 
With docActive.Shapes.AddLine(100, 100, 200, 300).Line 
 .BeginArrowheadLength = msoArrowheadShort 
 .BeginArrowheadStyle = msoArrowheadOval 
 .BeginArrowheadWidth = msoArrowheadNarrow 
 .EndArrowheadLength = msoArrowheadLong 
 .EndArrowheadStyle = msoArrowheadTriangle 
 .EndArrowheadWidth = msoArrowheadWide 
End With
```


## See also


[LineFormat Object](Word.LineFormat.md)

