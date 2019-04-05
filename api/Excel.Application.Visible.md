---
title: Application.Visible property (Excel)
keywords: vbaxl10.chm133229
f1_keywords:
- vbaxl10.chm133229
ms.prod: excel
api_name:
- Excel.Application.Visible
ms.assetid: 4d702074-7d76-7b43-25e1-11d6a440392f
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.Visible property (Excel)

Returns or sets a **Boolean** value that determines whether the object is visible. Read/write.


## Syntax

_expression_.**Visible**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.

## Example

```vb
'When used in a workbook this makes Excel invisible.
Application.Visible = False

'Waiting  five seconds, then showing Excel again.
Application.Wait Now + TimeValue("00:00:05")

'Makes Excel visible again.
Application.Visible = True

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
