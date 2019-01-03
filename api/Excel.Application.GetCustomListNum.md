---
title: Application.GetCustomListNum method (Excel)
keywords: vbaxl10.chm133141
f1_keywords:
- vbaxl10.chm133141
ms.prod: excel
api_name:
- Excel.Application.GetCustomListNum
ms.assetid: c4a97a96-333a-1021-7324-5cca4f0d9f3c
ms.date: 06/08/2017
---


# Application.GetCustomListNum method (Excel)

Returns the custom list number for an array of strings. You can use this method to match both built-in lists and custom-defined lists.


## Syntax

_expression_. `GetCustomListNum`( `_ListArray_` )

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ListArray_|Required| **Variant**|An array of strings.|

## Return value

Long


## Remarks

This method generates an error if there's no corresponding list.


## Example

This example deletes a custom list.


```vb
n = Application.GetCustomListNum(Array("cogs", "sprockets", _ 
 "widgets", "gizmos")) 
Application.DeleteCustomList n
```


## See also


[Application Object](Excel.Application(object).md)

