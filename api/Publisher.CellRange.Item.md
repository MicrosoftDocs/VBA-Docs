---
title: CellRange.Item method (Publisher)
keywords: vbapb10.chm5177344
f1_keywords:
- vbapb10.chm5177344
ms.prod: publisher
api_name:
- Publisher.CellRange.Item
ms.assetid: 8f1fe143-e00c-7112-45dd-52158153cf28
ms.date: 06/06/2019
localization_priority: Normal
---


# CellRange.Item method (Publisher)

Returns an individual **[Cell](Publisher.Cell.md)** object in the specified **CellRange** collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[CellRange](Publisher.CellRange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_|Required| **Long**|The number of the object to return.|

## Return value

Cell


## Example

This example returns the first cell from a **CellRange** collection.

```vb
Dim cllTemp As Cell 
 
Set cllTemp = ActiveDocument.Pages(Index:=1).Shapes(1).Table.Cells.Item(Index:=1)
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]