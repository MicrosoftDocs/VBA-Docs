---
title: WorksheetFunction.Trim method (Excel)
keywords: vbaxl10.chm137126
f1_keywords:
- vbaxl10.chm137126
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Trim
ms.assetid: 1e596960-90d8-87f8-9f1f-3a5c9e302e0c
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.Trim method (Excel)

Removes all spaces from text except for single spaces between words. Use **Trim** on text that you have received from another application that may have irregular spacing.


## Syntax

_expression_.**Trim** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **String**|Text - the text from which you want spaces removed.|

## Return value

**String**


## Remarks

> [!IMPORTANT]
> The TRIM function in Excel was designed to trim the 7-bit ASCII space character (value 32) from text. In the Unicode character set, there is an additional space character called the nonbreaking space character that has a decimal value of 160. This character is commonly used in webpages as the HTML entity, `&nbsp;`. By itself, the **Trim** function and **WorksheetFunction.Trim** method do not remove this nonbreaking space character.

The **WorksheetFunction.Trim** method in Excel differs from the **[Trim](../Language/Reference/User-Interface-Help/ltrim-rtrim-and-trim-functions.md)** function in VBA, which removes only leading and trailing spaces.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]