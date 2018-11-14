---
title: LineFormat.EndArrowheadWidth property (Excel)
keywords: vbaxl10.chm110009
f1_keywords:
- vbaxl10.chm110009
ms.prod: excel
api_name:
- Excel.LineFormat.EndArrowheadWidth
ms.assetid: 12148fae-ede6-9b05-9283-710f2bb68bbf
ms.date: 06/08/2017
---


# LineFormat.EndArrowheadWidth property (Excel)

Returns or sets the width of the arrowhead at the end of the specified line. Read/write  **[MsoArrowheadWidth](Office.MsoArrowheadWidth.md)** .


## Syntax

 _expression_. `EndArrowheadWidth`

 _expression_ A variable that represents a [LineFormat](Excel.LineFormat.md) object.


## Remarks





| **MsoArrowheadWidth** can be one of these **MsoArrowheadWidth** constants.|
| **msoArrowheadNarrow**|
| **msoArrowheadWidthMedium**|
| **msoArrowheadWide**|
| **msoArrowheadWidthMixed**|

## Example

This example adds a line to  `myDocument`. There's a short, narrow oval on the line's starting point and a long, wide triangle on its end point.


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


## See also


[LineFormat Object](Excel.LineFormat.md)

