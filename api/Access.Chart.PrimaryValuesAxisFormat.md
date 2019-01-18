---
title: Chart.PrimaryValuesAxisFormat property (Access)
keywords: vbaac10.chm6164
f1_keywords:
- vbaac10.chm6164
ms.prod: access
api_name:
- Access.Chart.PrimaryValuesAxisFormat
ms.date: 11/28/2018
localization_priority: Normal
---


# Chart.PrimaryValuesAxisFormat property (Access)

Returns or sets the format of the values on the primary values axis. Read/write **String**.

You can use a [predefined or custom format](Access.format.propertynumber.and.currency.md).


## Syntax

_expression_.**PrimaryValuesAxisFormat**

_expression_ A variable that represents a **[Chart](Access.Chart.md)** object.


## Example

```vb
With myChart
 .PrimaryValuesAxisFormat = "#,###.#0"
 .SecondaryValuesAxisFormat = "Currency"
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]