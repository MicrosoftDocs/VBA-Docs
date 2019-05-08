---
title: Axis.HasTitle property (Excel)
keywords: vbaxl10.chm561083
f1_keywords:
- vbaxl10.chm561083
ms.prod: excel
api_name:
- Excel.Axis.HasTitle
ms.assetid: 4b3d656f-4416-42a6-cefd-9684ba98c8e3
ms.date: 04/13/2019
localization_priority: Normal
---


# Axis.HasTitle property (Excel)

**True** if the axis or chart has a visible title. Read/write **Boolean**.


## Syntax

_expression_.**HasTitle**

_expression_ A variable that represents an **[Axis](Excel.Axis(object).md)** object.


## Remarks

An axis title is represented by an **[AxisTitle](Excel.AxisTitle(object).md)** object.


## Example

This example adds an axis label to the category axis on Chart1.

```vb
With Charts("Chart1").Axis(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Text = "July Sales" 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]