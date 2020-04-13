---
title: InlineShapes.AddOLEControl method (Word)
keywords: vbawd10.chm162070630
f1_keywords:
- vbawd10.chm162070630
ms.prod: word
api_name:
- Word.InlineShapes.AddOLEControl
ms.assetid: 390f1a37-163f-42f7-5784-9730aa79e1d9
ms.date: 06/08/2017
localization_priority: Normal
---


# InlineShapes.AddOLEControl method (Word)

Creates an ActiveX control (formerly known as an OLE control). Returns the **[InlineShape](Word.InlineShape.md)** object that represents the new ActiveX control.


## Syntax

_expression_. `AddOLEControl`( `_ClassType_` , `_Range_` )

_expression_ Required. A variable that represents an '[InlineShapes](Word.inlineshapes.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ClassType_|Optional| **Variant**|The programmatic identifier for the ActiveX control to be created.|
| _Range_|Optional| **Variant**|The range where the ActiveX control will be placed in the text. The ActiveX control replaces the range, if the range isn't collapsed. If this argument is omitted, the Active X control is placed automatically.|

## Remarks

ActiveX controls are represented as either  **Shape** objects or **[InlineShape](Word.InlineShape.md)** objects in Microsoft Word. To modify the properties for an ActiveX control, you use the **Object** property of the **OLEFormat** object for the specified shape or inline shape.



For information about available ActiveX control class types, see [OLE Programmatic Identifiers](overview/Word.md).


## See also


[InlineShapes Collection Object](Word.inlineshapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]