---
title: Index.SortBy property (Word)
keywords: vbawd10.chm159186952
f1_keywords:
- vbawd10.chm159186952
ms.prod: word
api_name:
- Word.Index.SortBy
ms.assetid: 384e1d3c-5cfd-240d-95dd-fc8b7bc99283
ms.date: 06/08/2017
localization_priority: Normal
---


# Index.SortBy property (Word)

Returns or sets the sorting criteria for the specified index. Read/write  **WdIndexSortBy**.


## Syntax

 _expression_. `SortBy`

 _expression_ Required. A variable that represents an '[Index](Word.Index.md)' object.


## Example

This example sets the first index in the current document to sort by the number of strokes.


```vb
ActiveDocument.Indexes(1).SortBy = _ 
 wdIndexSortByStroke
```


## See also


[Index Object](Word.Index.md)

