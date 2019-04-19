---
title: WorksheetFunction.FilterXML method (Excel)
keywords: vbaxl10.chm137465
f1_keywords:
- vbaxl10.chm137465
ms.prod: excel
ms.assetid: bcaa41a9-a122-ee87-29ca-cabb224358a1
ms.date: 06/08/2017
localization_priority: Normal
---


# WorksheetFunction.FilterXML method (Excel)

Get specific data from the returned XML, typically from a  **WebService** function call.


## Syntax

_expression_. `FilterXML`_(Arg1,_ _Arg2)_

_expression_ A variable that represents a object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required|**String**|Valid xml string.|
| _Arg2_|Required|**String**|XPath query string.|

### Remarks

The XPath parameter is limited to 1024 characters.

The  **FILTERXML** function returns results that are parsed via the user specified data locale.


## Return value

 **VARIANT**


## See also


[WorksheetFunction Object](Excel.WorksheetFunction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]