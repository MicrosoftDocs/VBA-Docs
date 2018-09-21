---
title: SortFields.Add2 Method (Excel)
keywords: vbaxl10.chm152090
f1_keywords:
- vbaxl10.chm152090
ms.prod: excel
api_name:
- Excel.SortFields.Add2
ms.assetid: 
ms.date: 09/21/2018
---


# SortFields.Add Method (Excel)

Creates a new sort field and returns a  **SortFields** object.


## Syntax

 _expression_. `Add`( `_Key_` , `_SortOn_` , `_Order_` , `_CustomOrder_` , `_DataOption_` )

 _expression_ A variable that represents a [SortFields](./Excel.SortFields.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Key_|Required| **Range**|Specifies a key value for the sort.|
| _SortOn_|Optional| **Variant**|The field to sort on.|
| _Order_|Optional| **Variant**|Specifies the sort order.|
| _CustomOrder_|Optional| **Variant**|Specifies if a custom sort order should be used.|
| _DataOption_|Optional| **Variant**|Specifies the data option.|

### Return Value

SortField


## See also


[SortFields Object](Excel.SortFields.md)

