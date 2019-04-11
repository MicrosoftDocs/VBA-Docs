---
title: Visible property (Excel Graph)
keywords: vbagr10.chm66094
f1_keywords:
- vbagr10.chm66094
ms.prod: excel
ms.assetid: 8a2b1b7a-b880-0e43-ca9f-c5d2207f7cfd
ms.date: 04/12/2019
localization_priority: Normal
---


# Visible property (Excel Graph)

The **Visible** property as it applies to the following objects.

## Application object

Determines whether the object is visible. Read/write **Boolean**.

### Syntax

_expression_.**Visible**

_expression_ Required. An expression that returns an **[Application](excel.application-graph-object.md)** object.


## ChartFillFormat object

Determines whether the application is visible. Read/write **[MsoTriState](office.msotristate.md)**.

### Syntax

_expression_.**Visible**

_expression_ Required. An expression that returns a **[ChartFillFormat](Excel.ChartFillFormat.md)** object.

### Example

This example formats the chart's fill with a preset gradient and then makes the fill visible.

```vb
With myChart.ChartArea.Fill 
 .Visible = msoTrue 
 .PresetGradient msoGradientDiagonalDown, _ 
 3, msoGradientBrass 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]