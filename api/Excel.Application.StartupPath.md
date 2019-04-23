---
title: Application.StartupPath property (Excel)
keywords: vbaxl10.chm133212
f1_keywords:
- vbaxl10.chm133212
ms.prod: excel
api_name:
- Excel.Application.StartupPath
ms.assetid: 04bdd294-8127-37f2-7a39-b42923ac45b5
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.StartupPath property (Excel)

Returns the complete path of the startup folder, excluding the final separator. Read-only **String**.


## Syntax

_expression_.**StartupPath**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example displays the full path to the Microsoft Excel startup folder.

```vb
MsgBox Application.StartupPath
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]