---
title: Application.CommandUnderlines property (Excel)
keywords: vbaxl10.chm133095
f1_keywords:
- vbaxl10.chm133095
ms.prod: excel
api_name:
- Excel.Application.CommandUnderlines
ms.assetid: 07d3ea82-6ef4-db6f-f3cf-bef992664408
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.CommandUnderlines property (Excel)

Returns or sets the state of the command underlines in Microsoft Excel for the Macintosh. Can be one of the constants of **[XlCommandUnderlines](Excel.XlCommandUnderlines.md)**. Read/write **Long**.


## Syntax

_expression_.**CommandUnderlines**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

In Microsoft Excel for Windows, reading this property always returns **xlCommandUnderlinesOn**, and setting this property to anything other than **xlCommandUnderlinesOn** is an error.


## Example

This example turns off command underlines in Microsoft Excel for the Macintosh.

```vb
Application.CommandUnderlines = xlCommandUnderlinesOff
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]