---
title: WorksheetFunction.EncodeUrl method (Excel)
keywords: vbaxl10.chm137467
f1_keywords:
- vbaxl10.chm137467
ms.prod: excel
ms.assetid: f98a7c18-46fe-4a3b-93ad-78c6a6e06061
ms.date: 06/08/2017
localization_priority: Normal
---


# WorksheetFunction.EncodeUrl method (Excel)

URL encodes the argument.


## Syntax

_expression_. `EncodeUrl`_(Arg1)_

_expression_ A variable that represents a [WorksheetFunction](Excel.WorksheetFunction.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required|**String**|Text to be encoded.|

## Return value

 **VARIANT**


### Remarks

This method enables the referencing of other cells as arguments into the Web Service URL, as this will ensure the data will be URL encoded.


## See also


[WorksheetFunction Object](Excel.WorksheetFunction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]