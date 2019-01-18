---
title: WorksheetFunction.WebService method (Excel)
keywords: vbaxl10.chm137466
f1_keywords:
- vbaxl10.chm137466
ms.prod: excel
ms.assetid: 7b7be122-2b68-0431-6687-cc5dad897274
ms.date: 06/08/2017
localization_priority: Normal
---


# WorksheetFunction.WebService method (Excel)

Underlying function that calls the web service asynchronously, using an HTTP GET request, and returns the response.


## Syntax

_expression_. `WebService`_(Arg1)_

_expression_ A variable that represents a [WorksheetFunction object (Excel)](Excel.WorksheetFunction.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required|STRING|The URL of the web service to make the HTTP GET request to.|

### Remarks

The XPath parameter is limited to 1024 characters.


## Return value

 **VARIANT**


## See also


[WorksheetFunction Object](Excel.WorksheetFunction.md)

