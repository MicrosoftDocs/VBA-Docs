---
title: Fields.Item method (Publisher)
keywords: vbapb10.chm6029312
f1_keywords:
- vbapb10.chm6029312
ms.prod: publisher
api_name:
- Publisher.Fields.Item
ms.assetid: 95783e5a-2c82-235e-75a4-5ac15938718e
ms.date: 06/07/2019
localization_priority: Normal
---


# Fields.Item method (Publisher)

Returns an individual object in a specified collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[Fields](Publisher.Fields.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_|Required| **Long**|The number of the object to return.|

## Return value

Field


## Example

This example returns the first field from a **Fields** collection.

```vb
Dim fldTemp As Field 
 
Set fldTemp = ActiveDocument.Pages(Index:=1) _ 
 .Shapes(1).TextFrame.TextRange.Fields.Item(Index:=1)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]