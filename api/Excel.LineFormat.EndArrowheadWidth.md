---
title: LineFormat.EndArrowheadWidth property (Excel)
keywords: vbaxl10.chm110009
f1_keywords:
- vbaxl10.chm110009
ms.prod: excel
api_name:
- Excel.LineFormat.EndArrowheadWidth
ms.assetid: 12148fae-ede6-9b05-9283-710f2bb68bbf
ms.date: 04/30/2019
localization_priority: Normal
---


# LineFormat.EndArrowheadWidth property (Excel)

Returns or sets the width of the arrowhead at the end of the specified line. Read/write **[MsoArrowheadWidth](Office.MsoArrowheadWidth.md)**.


## Syntax

_expression_.**EndArrowheadWidth**

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