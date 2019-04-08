---
title: AutoText property (Excel Graph)
keywords: vbagr10.chm65671
f1_keywords:
- vbagr10.chm65671
ms.prod: excel
api_name:
- Excel.AutoText
ms.assetid: 629627fc-f7b9-b7e9-1675-195bfb435b54
ms.date: 04/09/2019
localization_priority: Normal
---


# AutoText property (Excel Graph)

**True** if the object automatically generates appropriate text based on context. Read/write **Boolean**.

## Syntax

_expression_.**AutoText**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the data labels for series one to automatically generate appropriate text.

```vb
myChart.SeriesCollection(1).DataLabels.AutoText = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]