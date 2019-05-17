---
title: CommentsThreaded.Item method (Excel)
keywords: vbaxl10.chm1008074
f1_keywords:
- vbaxl10.chm1008074
ms.prod: excel
api_name:
- Excel.CommentsThreaded.Item
ms.date: 05/15/2019
localization_priority: Normal
---


# CommentsThreaded.Item method (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[CommentsThreaded](Excel.CommentsThreaded.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number for the object.|

## Return value

A **[CommentThreaded](Excel.CommentThreaded.md)** object contained by the collection.


## Example

This example updates the text of threaded comment two.

```vb
Worksheets(1).CommentsThreaded.Item(2).Text "Updated Comment"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]