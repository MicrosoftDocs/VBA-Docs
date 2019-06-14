---
title: Shapes.AddTable method (Publisher)
keywords: vbapb10.chm2162713
f1_keywords:
- vbapb10.chm2162713
ms.prod: publisher
api_name:
- Publisher.Shapes.AddTable
ms.assetid: 1aa00f40-de41-12ed-8d4f-5e9c91cbf5af
ms.date: 06/14/2019
localization_priority: Normal
---


# Shapes.AddTable method (Publisher)

Adds a new **[Shape](Publisher.Shape.md)** object representing a table to the specified **Shapes** collection.


## Syntax

_expression_.**AddTable** (_NumRows_, _NumColumns_, _Left_, _Top_, _Width_, _Height_, _FixedSize_, _Direction_)

_expression_ A variable that represents a **[Shapes](Publisher.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_NumRows_|Required| **Long**|The number of rows in the new table. Values between 1 and 128 are valid; any values outside this range generate an error.|
|_NumColumns_|Required| **Long**|The number of columns in the new table. Values between 1 and 128 are valid; any values outside this range generate an error.|
|_Left_ |Required| **Variant**|The position of the left edge of the shape representing the table.|
|_Top_ |Required| **Variant**|The position of the top edge of the shape representing the table.|
|_Width_|Required| **Variant**|The width of the shape representing the table.|
|_Height_|Required| **Variant**|The height of the shape representing the table.|
|_FixedSize_|Optional| **Boolean**| **True** if Microsoft Publisher reduces the number of rows and columns of the table to fit the specified width and height. **False** if Publisher automatically increases the width and height of the table frame to accommodate the number of rows and columns in the table. Default is **False**.|
|_Direction_|Optional| **[PbTableDirectionType](publisher.pbtabledirectiontype.md)**|The direction in which table columns are numbered. The default depends on the current language setting.|

## Return value

Shape


## Remarks

For the _Left_, _Top_, _Width_, and _Height_ arguments, numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Publisher (for example, "2.5 in").

The _Direction_ parameter can be one of the **PbTableDirectionType** constants declared in the Microsoft Publisher type library and shown in the following table.

|Constant|Description|
|:-----|:-----|
| **pbTableDirectionLeftToRight**|Table columns are numbered from left to right. Default for left-to-right languages.|
| **pbTableDirectionRightToLeft**|Table columns are numbered from right to left. Default for right-to-left languages.|

## Example

This example creates a new table on the first page of the active publication.

```vb
Dim shpTable As Shape 
 
Set shpTable = ActiveDocument.Pages(1).Shapes.AddTable _ 
 (NumRows:=3, NumColumns:=4, _ 
 Left:=10, Top:=10, _ 
 Width:=288, Height:=216) 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]