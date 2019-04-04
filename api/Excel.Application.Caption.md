---
title: Application.Caption property (Excel)
keywords: vbaxl10.chm133088
f1_keywords:
- vbaxl10.chm133088
ms.prod: excel
api_name:
- Excel.Application.Caption
ms.assetid: 618f5623-2eb7-4b7e-2f15-c30a0c2e0fe2
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.Caption property (Excel)

Returns or sets a **String** value that represents the name that appears in the title bar of the main Microsoft Excel window.


## Syntax

_expression_.**Caption**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

If you don't set a name, or if you set the name to **Empty**, this property returns "Microsoft Excel."


## Example

This example sets the name that appears in the title bar of the main Microsoft Excel window to a custom name.


```vb
Application.Caption = "Blue Sky Airlines Reservation System"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
