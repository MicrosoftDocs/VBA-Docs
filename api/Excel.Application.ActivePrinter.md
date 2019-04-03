---
title: Application.ActivePrinter property (Excel)
keywords: vbaxl10.chm183078
f1_keywords:
- vbaxl10.chm183078
ms.prod: excel
api_name:
- Excel.Application.ActivePrinter
ms.assetid: 72c4a525-27ab-f109-64d3-bcc7a12df505
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.ActivePrinter property (Excel)

Returns or sets the name of the active printer. Read/write **String**.


## Syntax

_expression_.**ActivePrinter**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example displays the name of the active printer.

```vb
MsgBox "The name of the active printer is " & _ 
 Application.ActivePrinter
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
