---
title: Chart.ChartValues property (Access)
keywords: vbaac10.chm6108
f1_keywords:
- vbaac10.chm6108
ms.prod: access
api_name:
- Access.Chart.ChartValues
ms.date: 11/28/2018
localization_priority: Normal
---


# Chart.ChartValues property (Access)

Returns or sets the semicolon-separated list of field(s) used to determine the data series plotted on the value axis. Read/write **String**.


## Syntax

_expression_.**ChartValues**

_expression_ A variable that represents a **[Chart](Access.Chart.md)** object.


## Example

```vb
With myChart
 .ChartValues = "[Price];[Cost]"
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]