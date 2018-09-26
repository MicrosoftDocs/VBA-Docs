---
title: Hyperlinks.Item Property (Excel)
keywords: vbaxl10.chm534075
f1_keywords:
- vbaxl10.chm534075
ms.prod: excel
api_name:
- Excel.Hyperlinks.Item
ms.assetid: c3650cd1-1788-549a-e203-4d7bd6f049c2
ms.date: 06/08/2017
---


# Hyperlinks.Item Property (Excel)

Returns a single object from a collection.


## Syntax

 _expression_. `Item`( `_Index_` )

 _expression_ A variable that represents a [Hyperlinks](Excel.Hyperlinks.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number of the object.|

## Example

The following example activates hyperlink one on cell E5.


```vb
Worksheets(1).Range("E5").Hyperlinks.Item(1).Follow
```


## See also


[Hyperlinks Object](Excel.Hyperlinks.md)

