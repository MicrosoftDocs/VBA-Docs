---
title: LineFormat.EndArrowheadStyle property (Excel)
keywords: vbaxl10.chm110008
f1_keywords:
- vbaxl10.chm110008
ms.prod: excel
api_name:
- Excel.LineFormat.EndArrowheadStyle
ms.assetid: 0d9eaff5-3ebc-572c-e188-d39848fa9bd2
ms.date: 04/30/2019
localization_priority: Normal
---


# LineFormat.EndArrowheadStyle property (Excel)

Returns or sets the style of the arrowhead at the end of the specified line. Read/write **[MsoArrowheadStyle](Office.MsoArrowheadStyle.md)**.


## Syntax

_expression_.**EndArrowheadStyle**

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