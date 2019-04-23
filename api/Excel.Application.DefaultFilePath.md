---
title: Application.DefaultFilePath property (Excel)
keywords: vbaxl10.chm133115
f1_keywords:
- vbaxl10.chm133115
ms.prod: excel
api_name:
- Excel.Application.DefaultFilePath
ms.assetid: 8eb8f6a2-f5fe-0b7e-172f-e7cfabef4af2
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.DefaultFilePath property (Excel)

Returns or sets the default path that Microsoft Excel uses when it opens files. Read/write **String**.


## Syntax

_expression_.**DefaultFilePath**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example displays the current default file path.

```vb
MsgBox "The current default file path is " & _ 
 Application.DefaultFilePath
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]