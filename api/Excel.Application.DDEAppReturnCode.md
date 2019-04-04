---
title: Application.DDEAppReturnCode property (Excel)
keywords: vbaxl10.chm183089
f1_keywords:
- vbaxl10.chm183089
ms.prod: excel
api_name:
- Excel.Application.DDEAppReturnCode
ms.assetid: 9b55dcce-eea8-a8b7-dace-296191de18a4
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.DDEAppReturnCode property (Excel)

Returns the application-specific DDE return code that was contained in the last DDE acknowledge message received by Microsoft Excel. Read-only **Long**.


## Syntax

_expression_.**DDEAppReturnCode**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example sets the variable `appErrorCode` to the DDE return code.

```vb
appErrorCode = Application.DDEAppReturnCode
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]