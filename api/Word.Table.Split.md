---
title: Table.Split Method (Word)
keywords: vbawd10.chm156303378
f1_keywords:
- vbawd10.chm156303378
ms.prod: word
api_name:
- Word.Table.Split
ms.assetid: a96c6dff-8508-2a73-2f3a-fac755e026ff
ms.date: 06/08/2017
---


# Table.Split Method (Word)

Inserts an empty paragraph immediately above the specified row in the table, and returns a  **Table** object that contains both the specified row and the rows that follow it.


## Syntax

 _expression_. `Split`( `_BeforeRow_` )

 _expression_ Required. A variable that represents a '[Table](Word.Table.md)' object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _BeforeRow_|Required| **Variant**|The row that the table is to be split before. Can be a row number or a  **Row** object.|

### Return value

Table


## See also


[Table Object](Word.Table.md)

