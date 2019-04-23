---
title: Item method (VBA Add-In Object Model)
keywords: vbob6.chm104043
f1_keywords:
- vbob6.chm104043
ms.prod: office
ms.assetid: 46074a24-356c-f003-d8cd-67807bea1bcd
ms.date: 12/06/2018
localization_priority: Normal
---


# Item method (VBA Add-In Object Model)

Returns the indexed member of a [collection](../../Glossary/vbe-glossary.md#collection).

## Syntax

_object_.**Item** (_index_)

<br/>

The **Item** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.|
| _index_|Required. An expression that specifies the position of a member of the collection. If a [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression), _index_ must be a number from 1 to the value of the collection's **[Count](count-property-vba-add-in-object-model.md)** property. If a [string expression](../../Glossary/vbe-glossary.md#string-expression), _index_ must correspond to the _key_ [argument](../../Glossary/vbe-glossary.md#argument) specified when the member was added to the collection.|

<br/>

The following table lists the [collections](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md) and their corresponding _key_ arguments for use with the **Item** method. The string that you pass to the **Item** method must match the collection's _key_ argument.

|Collection|Key argument|
|:-----|:-----|
|**Windows**|**[Caption](caption-property-vba-add-in-object-model.md)** property setting|
|**LinkedWindows**|**Caption** property setting|
|**CodePanes**|No unique string is associated with this collection.|
|**VBProjects**|**[Name](name-property-vba-add-in-object-model.md)** property setting|
|**VBComponents**|**Name** property setting|
|**References**|**Name** property setting|
|**Properties**|**Name** property setting|

## Remarks

The _index_ argument can be a numeric value or a string containing the title of the object.

> [!IMPORTANT] 
> Objects, properties, and methods for controlling linked windows, linked window frames, and docked windows are included on the Macintosh for compatibility with code written in Windows. However, these language elements will generate run-time errors when run on the Macintosh.


## See also

- [Property object](../visual-basic-add-in-model/objects-visual-basic-add-in-model.md#property)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]