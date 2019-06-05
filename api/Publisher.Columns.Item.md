---
title: Columns.Item method (Publisher)
keywords: vbapb10.chm5046272
f1_keywords:
- vbapb10.chm5046272
ms.prod: publisher
api_name:
- Publisher.Columns.Item
ms.assetid: c16df25c-ea8d-c04e-bccd-7e642bb7198a
ms.date: 06/06/2019
localization_priority: Normal
---


# Columns.Item method (Publisher)

Returns an individual **[Column](Publisher.Column.md)** object in the specified **Columns** collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[Columns](Publisher.Columns.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_|Required| **Long**|The number of the object to return.|

## Return value

Column


## Example

This example returns the first column from a **Columns** collection.

```vb
Dim colTemp As Column 
 
Set colTemp = ActiveDocument.Pages(Index:=1) _ 
 .Shapes(1).Table.Columns.Item(Index:=1)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]