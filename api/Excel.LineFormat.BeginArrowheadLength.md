---
title: LineFormat.BeginArrowheadLength property (Excel)
keywords: vbaxl10.chm110003
f1_keywords:
- vbaxl10.chm110003
ms.prod: excel
api_name:
- Excel.LineFormat.BeginArrowheadLength
ms.assetid: 7116965a-601c-46b5-9cb6-6cd339cccb80
ms.date: 04/30/2019
localization_priority: Normal
---


# LineFormat.BeginArrowheadLength property (Excel)

Returns or sets the length of the arrowhead at the beginning of the specified line. Read/write **[MsoArrowheadLength](Office.MsoArrowheadLength.md)**.


## Syntax

_expression_.**BeginArrowheadLength**

_expression_ A variable that represents a **[LineFormat](Excel.LineFormat.md)** object.


## Example

This example adds a line to _myDocument_. There's a short, narrow oval on the line's starting point and a long, wide triangle on its end point.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddLine(100, 100, 200, 300).Line 
    .BeginArrowheadLength = msoArrowheadShort 
    .BeginArrowheadStyle = msoArrowheadOval 
    .BeginArrowheadWidth = msoArrowheadNarrow 
    .EndArrowheadLength = msoArrowheadLong 
    .EndArrowheadStyle = msoArrowheadTriangle 
    .EndArrowheadWidth = msoArrowheadWide 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]