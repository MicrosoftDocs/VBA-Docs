---
title: ShapeRange.Item Method (Publisher)
keywords: vbapb10.chm2293760
f1_keywords:
- vbapb10.chm2293760
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Item
ms.assetid: f316bbac-b0be-0281-585b-c32dcb709b66
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.Item Method (Publisher)

Returns an individual object in a specified collection.


## Syntax

 _expression_. **Item**(**_Index_**)

 _expression_ A variable that represents a  **ShapeRange** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|Index|Required| **Variant**|The number or name of the field or list box item to return.|

## Return value

Shape


## Example

This example returns the first shape inside a grouped shape.


```vb
Dim shpTemp As Shape 
 
Set shpTemp = ActiveDocument.Pages(Index:=1) _ 
 .Shapes(1).GroupItems.Item(Index:=1)
```


