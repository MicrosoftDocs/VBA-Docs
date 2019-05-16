---
title: SparklineGroups.Ungroup method (Excel)
keywords: vbaxl10.chm869081
f1_keywords:
- vbaxl10.chm869081
ms.prod: excel
api_name:
- Excel.SparklineGroups.Ungroup
ms.assetid: c67c54f4-d5d1-5f12-2413-671db612a954
ms.date: 05/16/2019
localization_priority: Normal
---


# SparklineGroups.Ungroup method (Excel)

Ungroups the sparklines in the selected sparkline group.


## Syntax

_expression_.**Ungroup**

_expression_ A variable that represents a **[SparklineGroups](Excel.SparklineGroups.md)** object.


## Return value

Nothing


## Example

The following code example selects the range A1:A4 and ungroups the sparklines in that range.

```vb
Range("A1:A4").Select 
Selection.SparklineGroups.Ungroup
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]