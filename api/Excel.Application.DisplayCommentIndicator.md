---
title: Application.DisplayCommentIndicator property (Excel)
keywords: vbaxl10.chm133123
f1_keywords:
- vbaxl10.chm133123
ms.prod: excel
api_name:
- Excel.Application.DisplayCommentIndicator
ms.assetid: 8617da4e-97cb-fe57-bb51-a9c671e2ff27
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.DisplayCommentIndicator property (Excel)

Returns or sets the way cells display comments and indicators. Can be one of the **[XlCommentDisplayMode](Excel.XlCommentDisplayMode.md)** constants.


## Syntax

_expression_.**DisplayCommentIndicator**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example hides cell tips but retains comment indicators.

```vb
Application.DisplayCommentIndicator = xlCommentIndicatorOnly
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]