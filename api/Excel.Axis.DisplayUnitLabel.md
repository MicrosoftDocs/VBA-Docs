---
title: Axis.DisplayUnitLabel property (Excel)
keywords: vbaxl10.chm561116
f1_keywords:
- vbaxl10.chm561116
ms.prod: excel
api_name:
- Excel.Axis.DisplayUnitLabel
ms.assetid: e3a78e7b-464e-80b0-8bde-49f08ab4c842
ms.date: 04/13/2019
localization_priority: Normal
---


# Axis.DisplayUnitLabel property (Excel)

Returns the **[DisplayUnitLabel](Excel.DisplayUnitLabel(object).md)** object for the specified axis. Returns **null** if the **[HasDisplayUnitLabel](Excel.Axis.HasDisplayUnitLabel.md)** property is set to **False**. Read-only.


## Syntax

_expression_.**DisplayUnitLabel**

_expression_ A variable that represents an **[Axis](Excel.Axis(object).md)** object.


## Example

This example sets the label caption to Millions for the value axis on Chart1, and then it turns off automatic font scaling.

```vb
With Charts("Chart1").Axes(xlValue).DisplayUnitLabel 
 .Caption = "Millions" 
 .AutoScaleFont = False 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]