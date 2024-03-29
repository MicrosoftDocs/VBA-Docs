---
title: Chart.Elevation property (Excel)
keywords: vbaxl10.chm149106
f1_keywords:
- vbaxl10.chm149106
api_name:
- Excel.Chart.Elevation
ms.assetid: 44dde783-5bf7-7c5c-475b-0666337249d7
ms.date: 04/16/2019
ms.localizationpriority: medium
---


# Chart.Elevation property (Excel)

Returns or sets the elevation of the 3D chart view, in degrees. Read/write **Long**.


## Syntax

_expression_.**Elevation**

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Remarks

The chart elevation is the height at which you view the chart, in degrees. The default is 15 for most chart types. The value of this property must be between -90 and 90, except for 3D bar charts, where it must be between 0 and 44.


## Example

This example sets the chart elevation of Chart1 to 34 degrees. The example should be run on a 3D chart (the **Elevation** property fails on 2D charts).

```vb
Charts("Chart1").Elevation = 34
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]