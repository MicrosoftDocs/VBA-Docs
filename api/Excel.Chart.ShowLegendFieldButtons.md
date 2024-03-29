---
title: Chart.ShowLegendFieldButtons property (Excel)
keywords: vbaxl10.chm149190
f1_keywords:
- vbaxl10.chm149190
api_name:
- Excel.Chart.ShowLegendFieldButtons
ms.assetid: 44f1554c-145b-8600-07c4-40b6891dab2d
ms.date: 04/16/2019
ms.localizationpriority: medium
---


# Chart.ShowLegendFieldButtons property (Excel)

Returns or sets whether to display legend field buttons on a PivotChart. Read/write.


## Syntax

_expression_.**ShowLegendFieldButtons**

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Return value

**Boolean**


## Remarks

Set the **ShowLegendFieldButtons** property to **True** to display legend field buttons on the specified PivotChart. Set the property to **False** to hide the buttons.

The **ShowLegendFieldButtons** property corresponds to the **Show Legend Field Buttons** command on the **Field Buttons** drop-down list of the **Analyze** tab, which is available when a PivotChart is selected.


## Example

The following code example sets Chart 1 to display legend field buttons.

```vb
ActiveSheet.ChartObjects("Chart 1").Activate 
ActiveChart.ShowLegendFieldButtons = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]