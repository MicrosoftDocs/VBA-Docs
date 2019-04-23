---
title: Indexes.AutoMarkEntries method (Word)
keywords: vbawd10.chm159121511
f1_keywords:
- vbawd10.chm159121511
ms.prod: word
api_name:
- Word.Indexes.AutoMarkEntries
ms.assetid: ff348374-58f4-1ae6-3d3d-4978924df571
ms.date: 06/08/2017
localization_priority: Normal
---


# Indexes.AutoMarkEntries method (Word)

Automatically adds XE (Index Entry) fields to the specified document, using the entries from a concordance file.


## Syntax

_expression_. `AutoMarkEntries`( `_ConcordanceFileName_` )

_expression_ Required. A variable that represents an '[Indexes](Word.indexes.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ConcordanceFileName_|Required| **String**|The concordance file name that includes a list of items to be indexed.|

## Remarks

A concordance file is a Word document that contains a two-column table, with terms to index in the first column and index entries in the second column.


## Example

This example adds index entries to Thesis.doc based on the entries in C:\Documents\List.doc.


```vb
Documents("Thesis.doc").Indexes.AutoMarkEntries _ 
 ConcordanceFileName:="C:\Documents\List.doc"
```


## See also


[Indexes Collection Object](Word.indexes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]