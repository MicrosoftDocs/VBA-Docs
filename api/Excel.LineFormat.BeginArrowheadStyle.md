---
title: LineFormat.BeginArrowheadStyle property (Excel)
keywords: vbaxl10.chm110004
f1_keywords:
- vbaxl10.chm110004
ms.prod: excel
api_name:
- Excel.LineFormat.BeginArrowheadStyle
ms.assetid: 5f327e3f-d6bf-9709-e6bb-7be7f701899b
ms.date: 04/30/2019
localization_priority: Normal
---


# LineFormat.BeginArrowheadStyle property (Excel)

Returns or sets the style of the arrowhead at the beginning of the specified line. Read/write **[MsoArrowheadStyle](Office.MsoArrowheadStyle.md)**.


## Syntax

_expression_.**BeginArrowheadStyle**

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