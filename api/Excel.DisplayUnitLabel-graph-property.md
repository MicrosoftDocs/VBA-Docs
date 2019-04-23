---
title: DisplayUnitLabel property (Excel Graph)
keywords: vbagr10.chm67318
f1_keywords:
- vbagr10.chm67318
ms.prod: excel
api_name:
- Excel.DisplayUnitLabel
ms.assetid: 50e91894-9b5d-c915-e94c-e4563b54487a
ms.date: 04/10/2019
localization_priority: Normal
---


# DisplayUnitLabel property (Excel Graph)

Returns the **DisplayUnitLabel** object for the value axis in the specified chart. Returns **Null** if the **[HasDisplayUnitLabel](Excel.HasDisplayUnitLabel.md)** property is **False**. Read-only.

## Syntax

_expression_.**DisplayUnitLabel**

_expression_ Required. An expression that returns a **[DisplayUnitLabel](Excel.DisplayUnitLabel-graph-object.md)** object.

## Example

This example sets the caption for the value axis in _myChart_ to Millions, and turns off automatic font scaling.

```vb
With myChart.Axes(xlValue).DisplayUnitLabel 
 .Caption = "Millions" 
 .AutoScaleFont = False 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]