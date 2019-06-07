---
title: InlineShapes.Range property (Publisher)
keywords: vbapb10.chm5767173
f1_keywords:
- vbapb10.chm5767173
ms.prod: publisher
api_name:
- Publisher.InlineShapes.Range
ms.assetid: 375843c1-5198-6981-2e7c-8abd1d0e9dff
ms.date: 06/08/2019
localization_priority: Normal
---


# InlineShapes.Range property (Publisher)

Returns a **[ShapeRange](Publisher.ShapeRange.md)** collection that represents the same set of inline shapes as the **InlineShapes** collection whose method was called. This allows for miscellaneous formatting of the contained shapes. Read-only.


## Syntax

_expression_.**Range** (_Index_)

_expression_ A variable that represents an **[InlineShapes](Publisher.InlineShapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_|Optional| **Long**|The index position of the inline shape within the **ShapeRange** collection.|

## Example

The following example searches through each shape on the first page of the publication, and for all inline shapes within each shape, finds the first inline shape within the range of inline shapes and flips it vertically.

```vb
Dim theShape As Shape 
Dim theShapes As Shapes 
 
Set theShapes = ActiveDocument.Pages(1).Shapes 
 
For Each theShape In theShapes 
 With theShape.TextFrame.TextRange 
 .InlineShapes.Range(1).Flip (msoFlipVertical) 
 End With 
Next
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]