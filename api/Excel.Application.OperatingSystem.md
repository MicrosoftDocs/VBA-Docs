---
title: Application.OperatingSystem property (Excel)
keywords: vbaxl10.chm133187
f1_keywords:
- vbaxl10.chm133187
ms.prod: excel
api_name:
- Excel.Application.OperatingSystem
ms.assetid: a36c5080-1d7e-a941-1bad-94f92522c7cf
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.OperatingSystem property (Excel)

Returns the name and version number of the current operating system. Read-only **String**.

For example:

- "Windows (32-bit) 4.00" or "Macintosh 7.00"

- "Windows (32-bit) NT 6.02" with Win8.1 (=6.02, 64-bit) and Excel 2013 (15.0.4631.1000, 32-bit)

- "Windows (64-bit) NT :.00" with Win10 (64-bit) and Excel 2016 (16.0.6326.1010, 64-bit)




## Syntax

_expression_.**OperatingSystem**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example displays the name of the operating system.

```vb
MsgBox "Microsoft Excel is using " & Application.OperatingSystem
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]