---
title: LineFormat.BeginArrowheadWidth property (Excel)
keywords: vbaxl10.chm110005
f1_keywords:
- vbaxl10.chm110005
ms.prod: excel
api_name:
- Excel.LineFormat.BeginArrowheadWidth
ms.assetid: 82d9b8fe-4aa5-3292-f792-c14332c2103d
ms.date: 04/30/2019
localization_priority: Normal
---


# LineFormat.BeginArrowheadWidth property (Excel)

Returns or sets the width of the arrowhead at the beginning of the specified line. Read/write **[MsoArrowheadWidth](Office.MsoArrowheadWidth.md)**.


## Syntax

_expression_.**BeginArrowheadWidth**

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