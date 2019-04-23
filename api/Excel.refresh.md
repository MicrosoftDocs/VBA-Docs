---
title: Refresh method (Excel Graph)
keywords: vbagr10.chm3077631
f1_keywords:
- vbagr10.chm3077631
ms.prod: excel
ms.assetid: 6bb2b3ee-413e-ad0d-1b94-770b21c9ebcc
ms.date: 04/09/2019
localization_priority: Normal
---


# Refresh method (Excel Graph)

Causes the specified chart to be redrawn immediately.

## Syntax

_expression_.**Refresh**

_expression_ Required. An expression that returns a **[Chart](Excel.Chart-graph-object.md)** object.


## Example

This example refreshes the first chart in the application. This example assumes that a chart exists in the application.

```vb
Sub RefeshChart() 
 
 Application.Charts(1).Refresh 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]