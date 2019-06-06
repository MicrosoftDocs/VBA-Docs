---
title: Columns.Add method (Publisher)
keywords: vbapb10.chm5046276
f1_keywords:
- vbapb10.chm5046276
ms.prod: publisher
api_name:
- Publisher.Columns.Add
ms.assetid: b3dfb892-6bda-d2c4-11f7-9bd29bf257aa
ms.date: 06/06/2019
localization_priority: Normal
---


# Columns.Add method (Publisher)

Adds a new **[Column](Publisher.Column.md)** object to the specified **Columns** collection and returns the new **Column** object.


## Syntax

_expression_.**Add** (_BeforeColumn_)

_expression_ A variable that represents a **[Columns](Publisher.Columns.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_BeforeColumn_|Optional| **Long**|The number of the column before which to insert the new column. If this argument is omitted, the new column is added after the existing columns. An error occurs if the value of this argument does not correspond to an existing column in the table.|

## Return value

Column


## Example

The following example adds a column before column three in the specified table.

```vb
Dim colNew As Column 
 
Set colNew = ActiveDocument.Pages(1).Shapes(1) _ 
 .Table.Columns.Add(BeforeColumn:=3)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]