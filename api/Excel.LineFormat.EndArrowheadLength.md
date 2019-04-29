---
title: LineFormat.EndArrowheadLength property (Excel)
keywords: vbaxl10.chm110007
f1_keywords:
- vbaxl10.chm110007
ms.prod: excel
api_name:
- Excel.LineFormat.EndArrowheadLength
ms.assetid: e6dd340b-9732-db7e-2efb-7003bca0aea6
ms.date: 04/30/2019
localization_priority: Normal
---


# LineFormat.EndArrowheadLength property (Excel)

Returns or sets the length of the arrowhead at the end of the specified line. Read/write **[MsoArrowheadLength](Office.MsoArrowheadLength.md)**.


## Syntax

_expression_.**EndArrowheadLength**

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