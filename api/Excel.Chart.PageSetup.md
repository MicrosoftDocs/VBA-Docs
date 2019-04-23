---
title: Chart.PageSetup property (Excel)
keywords: vbaxl10.chm148085
f1_keywords:
- vbaxl10.chm148085
ms.prod: excel
api_name:
- Excel.Chart.PageSetup
ms.assetid: 9a47bfd6-10b5-5f8e-86c2-e56c468de9d8
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.PageSetup property (Excel)

Returns a **[PageSetup](Excel.PageSetup.md)** object that contains all the page setup settings for the specified object. Read-only.


## Syntax

_expression_.**PageSetup**

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Example

This example sets the center header text for Chart1.

```vb
Charts("Chart1").PageSetup.CenterHeader = "December Sales"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]