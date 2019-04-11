---
title: Version property (Excel Graph)
keywords: vbagr10.chm5208125
f1_keywords:
- vbagr10.chm5208125
ms.prod: excel
api_name:
- Excel.Version
ms.assetid: 16be4008-4999-135e-dc53-b0212bbedac9
ms.date: 04/12/2019
localization_priority: Normal
---


# Version property (Excel Graph)

Returns the Graph version number. Read-only **String**.

## Syntax

_expression_.**Version**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Example

This example displays a message box that contains the Graph version number.

```vb
MsgBox "Welcome to Graph version " & _ 
 myChart.Application.Version
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]