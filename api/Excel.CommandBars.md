---
title: CommandBars property (Excel Graph)
keywords: vbagr10.chm66975
f1_keywords:
- vbagr10.chm66975
ms.prod: excel
api_name:
- Excel.CommandBars
ms.assetid: 70c5ec17-7ce0-fd21-4f2f-6601b189266e
ms.date: 04/10/2019
localization_priority: Normal
---


# CommandBars property (Excel Graph)

Returns a **CommandBars** object that represents the Graph command bars. Read-only **CommandBars** object.

## Syntax

_expression_.**CommandBars**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Example

This example deletes all custom command bars that aren't visible.

```vb
For Each bar In myChart.Application.CommandBars 
 If Not bar.BuiltIn And Not bar.Visible Then bar.Delete 
Next
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]