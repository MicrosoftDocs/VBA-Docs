---
title: Chart.ChartTitle property (Excel)
keywords: vbaxl10.chm149089
f1_keywords:
- vbaxl10.chm149089
ms.prod: excel
api_name:
- Excel.Chart.ChartTitle
ms.assetid: 3a083c1f-7a3f-3368-c547-297f0e5d26cb
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.ChartTitle property (Excel)

Returns a **[ChartTitle](Excel.ChartTitle(object).md)** object that represents the title of the specified chart. Read-only.


## Syntax

_expression_.**ChartTitle**

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Example

This example sets the text for the title of Chart1.

```vb
With Charts("Chart1") 
 .HasTitle = True 
 .ChartTitle.Text = "First Quarter Sales" 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
