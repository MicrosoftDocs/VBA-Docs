---
title: Application.TransitionMenuKey property (Excel)
keywords: vbaxl10.chm133218
f1_keywords:
- vbaxl10.chm133218
ms.prod: excel
api_name:
- Excel.Application.TransitionMenuKey
ms.assetid: 3ea5b071-1ba7-19e9-1d6d-00bf128466e2
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.TransitionMenuKey property (Excel)

Returns or sets the Microsoft Excel menu or help key, which is usually `/`. Read/write **String**.


## Syntax

_expression_.**TransitionMenuKey**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example sets the transition menu key to `/` (which is the default).

```vb
Application.TransitionMenuKey = "/"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]