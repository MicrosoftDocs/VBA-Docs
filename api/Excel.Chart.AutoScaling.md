---
title: Chart.AutoScaling property (Excel)
keywords: vbaxl10.chm149080
f1_keywords:
- vbaxl10.chm149080
ms.prod: excel
api_name:
- Excel.Chart.AutoScaling
ms.assetid: fecafb42-56fb-3c33-dc03-cb290b4a28df
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.AutoScaling property (Excel)

**True** if Microsoft Excel scales a 3D chart so that it's closer in size to the equivalent 2D chart. The **[RightAngleAxes](Excel.Chart.RightAngleAxes.md)** property must be **True**. Read/write **Boolean**.


## Syntax

_expression_.**AutoScaling**

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Example

This example automatically scales Chart1. The example should be run on a 3D chart.

```vb
With Charts("Chart1") 
 .RightAngleAxes = True 
 .AutoScaling = True 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]