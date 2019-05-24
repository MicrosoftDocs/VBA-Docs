---
title: WorksheetFunction.WebService method (Excel)
keywords: vbaxl10.chm137466
f1_keywords:
- vbaxl10.chm137466
ms.prod: excel
ms.assetid: 7b7be122-2b68-0431-6687-cc5dad897274
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.WebService method (Excel)

Underlying function that calls the web service asynchronously, using an HTTP GET request, and returns the response.


## Syntax

_expression_.**WebService** (_Url_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Url_|Required|**String**|The URL of the web service to make the HTTP GET request to.|

## Remarks

The URL parameter is limited to 2048 characters.


## Return value

**Variant**


## See also

- [WEBSERVICE() function](https://support.office.com/article/webservice-function-0546a35a-ecc6-4739-aed7-c0b7ce1562c4)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
