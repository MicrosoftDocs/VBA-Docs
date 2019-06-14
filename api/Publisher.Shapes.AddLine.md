---
title: Shapes.AddLine method (Publisher)
keywords: vbapb10.chm2162708
f1_keywords:
- vbapb10.chm2162708
ms.prod: publisher
api_name:
- Publisher.Shapes.AddLine
ms.assetid: 43df8878-5640-875f-06e0-37e1feb47b78
ms.date: 06/14/2019
localization_priority: Normal
---


# Shapes.AddLine method (Publisher)

Adds a new **[Shape](Publisher.Shape.md)** object representing a line to the specified **Shapes** collection.


## Syntax

_expression_.**AddLine** (_BeginX_, _BeginY_, _EndX_, _EndY_)

_expression_ A variable that represents a **[Shapes](Publisher.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_BeginX_|Required| **Variant**|The x-coordinate of the beginning point of the line.|
|_BeginY_|Required| **Variant**|The y-coordinate of the beginning point of the line.|
|_EndX_|Required| **Variant**|The x-coordinate of the ending point of the line.|
|_EndY_|Required| **Variant**|The y-coordinate of the ending point of the line.|

## Return value

Shape


## Remarks

For the _BeginX_, _BeginY_, _EndX_, and _EndY_ arguments, numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").


## Example

The following example adds a new line to the first page of the active publication.

```vb
Dim shpLine As Shape 
 
Set shpLine = ActiveDocument.Pages(1).Shapes.AddLine _ 
 (BeginX:=144, BeginY:=144, _ 
 EndX:=180, EndY:=72) 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]