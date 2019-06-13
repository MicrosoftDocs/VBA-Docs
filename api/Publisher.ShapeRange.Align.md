---
title: ShapeRange.Align method (Publisher)
keywords: vbapb10.chm2294016
f1_keywords:
- vbapb10.chm2294016
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Align
ms.assetid: ef522d47-3fc7-cfca-5b9a-44ff020f8b31
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.Align method (Publisher)

Aligns all the shapes in the specified **ShapeRange** object.


## Syntax

_expression_.**Align** (_AlignCmd_, _RelativeTo_)

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_AlignCmd_|Required| **[MsoAlignCmd](office.msoaligncmd.md)** |Specifies how the shapes are to be aligned.|
|_RelativeTo_|Required| **[MsoTriState](office.msotristate.md)** |Specifies whether shapes are aligned relative to the page or to one another.|

## Remarks

The _AlignCmd_ parameter can be one of the **MsoAlignCmd** constants declared in the Microsoft Office type library and shown in the following table.

|Constant|Description|
|:-----|:-----|
| **msoAlignBottoms**|Aligns shapes along their bottom edges. If _RelativeTo_ is **msoFalse**, the bottommost shape determines the line against which the other shapes are aligned.|
| **msoAlignCenters**|Aligns shapes on a vertical line through their centers. If _RelativeTo_ is **msoFalse**, shapes are aligned on a line halfway between the left- and rightmost shapes.|
| **msoAlignLefts**|Aligns shapes along their left edges. If _RelativeTo_ is **msoFalse**, the leftmost shape determines the line against which the other shapes are aligned.|
| **msoAlignMiddles**|Aligns shapes on a horizontal line through their centers. If _RelativeTo_ is **msoFalse**, shapes are aligned on a line halfway between the top- and bottommost shapes.|
| **msoAlignRights**| **msoAlignRights** Aligns shapes along their right edges. If _RelativeTo_ is **msoFalse**, the rightmost shape determines the line against which the other shapes are aligned.|
| **msoAlignTops**| Aligns shapes along their top edges. If _RelativeTo_ is **msoFalse**, the topmost shape determines the line against which the other shapes are aligned.|

The _RelativeTo_ parameter can be one of the **MsoTriState** constants. 

|Constant|Description|
|:-----|:-----|
| **msoFalse**|Aligns shapes relative to one another.|
| **msoTrue**|Aligns shapes relative to the page.|

If the _RelativeTo_ parameter is **msoFalse** and the shape range contains only one shape, an error occurs.


## Example

The following example aligns all the shapes on the first page of the active publication on a vertical line through their centers.

```vb
ActiveDocument.Pages(1).Shapes.Range.Align _ 
 AlignCmd:=msoAlignCenters, _ 
 RelativeTo:=msoTrue 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]