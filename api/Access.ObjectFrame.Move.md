---
title: ObjectFrame.Move method (Access)
keywords: vbaac10.chm11625
f1_keywords:
- vbaac10.chm11625
ms.prod: access
api_name:
- Access.ObjectFrame.Move
ms.assetid: 63b05ea6-d761-adfa-5aa6-25d16ae5ed3c
ms.date: 06/08/2017
localization_priority: Normal
---


# ObjectFrame.Move method (Access)

Moves the specified object to the coordinates specified by the argument values.


## Syntax

_expression_. `Move`( ` _Left_`, ` _Top_`, ` _Width_`, ` _Height_` )

_expression_ A variable that represents an [ObjectFrame](Access.ObjectFrame.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Left_|Required|**Variant**|The screen position in [twips](../language/glossary/vbe-glossary.md#twip) for the left edge of the object relative to the left edge of the Microsoft Access window.|
| _Top_|Optional|**Variant**|The screen position in [twips](../language/glossary/vbe-glossary.md#twip) for the top edge of the object relative to the top edge of the Microsoft Access window.|
| _Width_|Optional|**Variant**|The desired width in [twips](../language/glossary/vbe-glossary.md#twip) of the object.|
| _Height_|Optional|**Variant**|The desired height in [twips](../language/glossary/vbe-glossary.md#twip) of the object.|

## Remarks

Only the  _Left_ argument is required. However, to specify any other arguments, you must specify all the arguments that precede it. For example, you cannot specify _Width_ without specifying _Left_ and _Top_. Any trailing arguments that are unspecified remain unchanged.

This method overrides the  **Moveable** property.

In Datasheet View or Print Preview, changes made using the  **Move** method are saved if the user explicitly saves the database, but Access does not prompt the user to save such changes.


## See also


[ObjectFrame Object](Access.ObjectFrame.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]