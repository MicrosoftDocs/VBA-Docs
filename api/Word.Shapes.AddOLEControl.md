---
title: Shapes.AddOLEControl method (Word)
keywords: vbawd10.chm161415270
f1_keywords:
- vbawd10.chm161415270
ms.prod: word
api_name:
- Word.Shapes.AddOLEControl
ms.assetid: f0f5d8cb-ea31-58a9-f266-eff38610cf3b
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddOLEControl method (Word)

Creates an ActiveX control (formerly known as an OLE control). Returns the **InlineShape** object that represents the new ActiveX control.


## Syntax

_expression_. `AddOLEControl`( `_ClassType_` , `_Range_` )

_expression_ Required. A variable that represents a **[Shapes](Word.Shapes.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ClassType_|Optional| **Variant**|The programmatic identifier for the ActiveX control to be created.|
| _Range_|Optional| **Variant**|The range where the ActiveX control will be placed in the text. The ActiveX control replaces the range, if the range isn't collapsed. If this argument is omitted, the Active X control is placed automatically.|

## Return value

InlineShape


## Remarks

ActiveX controls are represented as either  **Shape** objects or **InlineShape** objects in Microsoft Word. To modify the properties for an ActiveX control, you use the **Object** property of the **OLEFormat** object for the specified shape or inline shape.

For information about available ActiveX control class types, see [OLE Programmatic Identifiers](overview/Word.md).


## See also


[Shapes Collection Object](Word.shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]