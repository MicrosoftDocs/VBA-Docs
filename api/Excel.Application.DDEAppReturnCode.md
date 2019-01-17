---
title: Application.DDEAppReturnCode property (Excel)
keywords: vbaxl10.chm183089
f1_keywords:
- vbaxl10.chm183089
ms.prod: excel
api_name:
- Excel.Application.DDEAppReturnCode
ms.assetid: 9b55dcce-eea8-a8b7-dace-296191de18a4
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DDEAppReturnCode property (Excel)

Returns the application-specific DDE return code that was contained in the last DDE acknowledge message received by Microsoft Excel. Read-only  **Long**.


## Syntax

_expression_. `DDEAppReturnCode`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Example

This example sets the variable  `appErrorCode` to the DDE return code.


```vb
appErrorCode = Application.DDEAppReturnCode
```


## See also


[Application Object](Excel.Application(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]