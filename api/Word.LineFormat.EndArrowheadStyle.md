---
title: LineFormat.EndArrowheadStyle property (Word)
keywords: vbawd10.chm164233322
f1_keywords:
- vbawd10.chm164233322
ms.prod: word
api_name:
- Word.LineFormat.EndArrowheadStyle
ms.assetid: 8893f334-4da7-ec32-f3e6-268706e3ca84
ms.date: 06/08/2017
localization_priority: Normal
---


# LineFormat.EndArrowheadStyle property (Word)

Returns or sets the style of the arrowhead at the end of the specified line. Read/write  **MsoArrowheadStyle**.


## Syntax

_expression_.**EndArrowheadStyle**

_expression_ Required. A variable that represents a **[LineFormat](Word.LineFormat.md)** object.


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

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]