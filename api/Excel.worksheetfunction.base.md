---
title: WorksheetFunction.Base method (Excel)
keywords: vbaxl10.chm137444
f1_keywords:
- vbaxl10.chm137444
ms.prod: excel
ms.assetid: df7544ca-e793-4fa8-a9a3-4f421b080723
ms.date: 06/08/2017
localization_priority: Normal
---


# WorksheetFunction.Base method (Excel)

Converts a number into a text representation with the given radix (base).


## Syntax

_expression_. `Base`_(Arg1,_ _Arg2,_ _Arg3)_

_expression_ A variable that represents a [WorksheetFunction](Excel.WorksheetFunction.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required|DOUBLE|The number that you want to convert.|
| _Arg2_|Required|DOUBLE|The base Radix that you want to convert the number into.|
| _Arg3_|Optional|**Variant**|The minimum length of the returned string. If omitted leading zeros are not added.|

## Return value

 **STRING**


## See also


[WorksheetFunction Object](Excel.WorksheetFunction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]