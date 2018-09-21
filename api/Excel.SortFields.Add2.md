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

Creates a new sort field and returns a  **SortFields** object which can optionally sort Data Types with the SubField defined.


## Syntax

 _expression_. `Add2`( `_Key_` , `_SortOn_` , `_Order_` , `_CustomOrder_` , `_DataOption_`, `_SubField_` )

 _expression_ A variable that represents a [SortFields](./Excel.SortFields.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Key_|Required| **Range**|Specifies a key value for the sort.|
| _SortOn_|Optional| **Variant**|The field to sort on.|
| _Order_|Optional| **Variant**|Specifies the sort order.|
| _CustomOrder_|Optional| **Variant**|Specifies if a custom sort order should be used.|
| _DataOption_|Optional| **Variant**|Specifies the data option.|
| _SubField_|Optional| **Variant**|Specifies the Field to sort on for a Data Type (Such as Population for Geography or Volume for Stocks).|

### Return Value

SortField


## See also


[SortFields Object](Excel.SortFields.md)

