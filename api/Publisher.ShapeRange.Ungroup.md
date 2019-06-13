---
title: ShapeRange.Ungroup method (Publisher)
keywords: vbapb10.chm2293801
f1_keywords:
- vbapb10.chm2293801
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Ungroup
ms.assetid: 253a366c-7317-14e7-2668-191eccec6cb8
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.Ungroup method (Publisher)

Ungroups the specified group of shapes or any groups of shapes in the specified shape range. If the specified shape is a picture or OLE object, Microsoft Publisher breaks it apart and converts it to an ungrouped set of shapes. For example, an embedded Microsoft Excel spreadsheet is converted into lines and text boxes. 

Returns the ungrouped shapes as a single **ShapeRange** object.


## Syntax

_expression_.**Ungroup**

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Return value

ShapeRange


## Remarks

Using this method on an inline shape or a shape that isn't a group, picture, or OLE object generates an error. Also, an error occurs if the picture is a bitmap, JPEG, GIF, or PNG (Portable Network Graphics) file.

Because a group of shapes is treated as a single object, grouping and ungrouping shapes changes the number of items in the **Shapes** collection and changes the index numbers of items that come after the affected items in the collection. 

Also, newly ungrouped shapes are added to the **Shapes** collection on the current page (or pages) or scratch area. As a result, they may shift from one collection to another.


## Example

This example ungroups any grouped shapes on the first page of the active publication.

```vb
Dim shpLoop As Shape 
 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 If shpLoop.Type = pbGroup Then shpLoop.Ungroup 
Next shpLoop 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]