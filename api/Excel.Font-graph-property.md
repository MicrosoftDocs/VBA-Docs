---
title: Font property (Excel Graph)
keywords: vbagr10.chm65682
f1_keywords:
- vbagr10.chm65682
api_name:
- Excel.Font
ms.assetid: 0bc46ec4-998e-043e-0713-9a381ec2b6ad
ms.date: 04/10/2019
ms.localizationpriority: medium
---


# Font property (Excel Graph)

Returns a **Font** object that represents the font of the specified object. Read/write **Font** object only for the **DataSheet** object; for all other objects, read-only **Font** object.

## Syntax

_expression_.**Font**

_expression_ Required. An expression that returns a **[Font](excel.font-graph-object.md)** object. 


## Example

This example sets the font in the chart title to 14-point bold italic.

```vb
With myChart.ChartTitle.Font 
 .Size = 14 
 .Bold = True 
 .Italic = True 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
