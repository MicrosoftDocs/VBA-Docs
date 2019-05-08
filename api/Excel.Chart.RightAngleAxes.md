---
title: Chart.RightAngleAxes property (Excel)
keywords: vbaxl10.chm149138
f1_keywords:
- vbaxl10.chm149138
ms.prod: excel
api_name:
- Excel.Chart.RightAngleAxes
ms.assetid: 632aa454-4113-97d3-a80c-eb745a950c6f
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.RightAngleAxes property (Excel)

**True** if the chart axes are at right angles, independent of chart rotation or elevation. Applies only to 3D line, column, and bar charts. Read/write **Boolean**.


## Syntax

_expression_.**RightAngleAxes**

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Remarks

If this property is **True**, the **[Perspective](Excel.Chart.Perspective.md)** property is ignored.


## Example

This example sets the axes on Chart1 to intersect at right angles. The example should be run on a 3D chart.

```vb
Charts("Chart1").RightAngleAxes = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]