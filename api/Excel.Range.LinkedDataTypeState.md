---
title: Range.LinkedDataTypeState property (Excel)
keywords: vbaxl10.chm144264
f1_keywords:
- vbaxl10.chm144264
ms.prod: excel
api_name:
- Excel.Range.LinkedDataTypeState
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.LinkedDataTypeState property (Excel)

Returns information about the state of any Linked data types, such as [Stocks or Geography](https://support.office.com/article/stock-quotes-and-geographic-data-61a33056-9935-484f-8ac8-f1a89e210877), in the range. Possible values are from the **[XlLinkedDataTypeState](Excel.XlLinkedDataTypeState.md)** enumeration. Read-only.


## Syntax

_expression_.**LinkedDataTypeState**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

For ranges that contains cells in different states, it will return **null**.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]