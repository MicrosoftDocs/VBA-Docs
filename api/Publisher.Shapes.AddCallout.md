---
title: Shapes.AddCallout method (Publisher)
keywords: vbapb10.chm2162704
f1_keywords:
- vbapb10.chm2162704
ms.prod: publisher
api_name:
- Publisher.Shapes.AddCallout
ms.assetid: bbf5f913-fcf0-b700-0c7e-9f0bdc7c6aea
ms.date: 06/14/2019
localization_priority: Normal
---


# Shapes.AddCallout method (Publisher)

Adds a new **[Shape](Publisher.Shape.md)** object representing a borderless line callout to the specified **Shapes** collection.


## Syntax

_expression_.**AddCallout** (_Type_, _Left_, _Top_, _Width_, _Height_)

_expression_ A variable that represents a **[Shapes](Publisher.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Type_|Required| **[MsoCalloutType](office.msocallouttype.md)**|The type of callout line. Can be one of the **MsoCalloutType** constants.|
|_Left_|Required| **Variant**|The position of the left edge of the shape representing the line callout.|
|_Top_|Required| **Variant**|The position of the top edge of the shape representing the line callout.|
|_Width_|Required| **Variant**|The width of the shape representing the line callout.|
|_Height_|Required| **Variant**|The height of the shape representing the line callout.|

## Return value

Shape


## Remarks

For the _Left_, _Top_, _Width_, and _Height_ arguments, numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").


## Example

The following example adds a new freely-rotating callout line to the first page of the active publication.

```vb
Dim shpCallout As Shape 
 
Set shpCallout = ActiveDocument.Pages(1).Shapes.AddCallout _ 
 (Type:=msoCalloutTwo, _ 
 Left:=144, Top:=216, _ 
 Width:=36, Height:=72)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]