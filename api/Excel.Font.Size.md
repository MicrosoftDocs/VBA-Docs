---
title: Font.Size property (Excel)
keywords: vbaxl10.chm559082
f1_keywords:
- vbaxl10.chm559082
ms.prod: excel
api_name:
- Excel.Font.Size
ms.assetid: 45f409cd-768b-0794-4fe9-ef002fa69606
ms.date: 06/08/2017
localization_priority: Priority
---


# Font.Size property (Excel)

Returns or sets the size of the font. Read/write **Variant**.

## Syntax

_expression_. `Size`

_expression_ A variable that represents a [Font](Excel.Font-graph-property.md) object.

## Example

This example sets the font size for cells A1:D10 on Sheet1 to 12 points.

```vb
With Worksheets("Sheet1").Range("A1:D10") 
 .Value = "Test" 
 .Font.Size = 12 
End With 

```

## See also

- [Font object](Excel.Font(object).md)

