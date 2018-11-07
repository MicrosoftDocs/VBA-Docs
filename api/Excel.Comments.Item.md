---
title: Comments.Item Method (Excel)
keywords: vbaxl10.chm514074
f1_keywords:
- vbaxl10.chm514074
ms.prod: excel
api_name:
- Excel.Comments.Item
ms.assetid: 87f0ecf0-9261-ffaf-39ca-4cdbc5712368
ms.date: 06/08/2017
---


# Comments.Item Method (Excel)

Returns a single object from a collection.


## Syntax

 _expression_. `Item`( `_Index_` )

 _expression_ A variable that represents a [Comments](Excel.Comments.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number for the object.|

## Return value

A  **[Comment](Excel.Comment.md)** object contained by the collection.


## Example

This example hides comment two.


```vb
Worksheets(1).Comments.Item(2).Visible = False
```


## See also


[Comments Object](Excel.Comments.md)

