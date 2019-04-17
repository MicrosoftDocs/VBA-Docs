---
title: Remove method (VBA Add-In Object Model)
keywords: vbob6.chm100142
f1_keywords:
- vbob6.chm100142
ms.prod: office
ms.assetid: acc163b9-e5ad-ef39-013a-614fc24bcde1
ms.date: 12/06/2018
localization_priority: Normal
---


# Remove method (VBA Add-In Object Model)

Removes an item from a [collection](../../Glossary/vbe-glossary.md#collection).

## Syntax

_object_.**Remove** (_component_)

<br/>

The **Remove** syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.|
| _component_|Required. For the **LinkedWindows** collection, an object.<br/><br/>For the **References** collection, a reference to a [type library](../../Glossary/vbe-glossary.md#type-library) or a [project](../../Glossary/vbe-glossary.md#project).<br/><br/>For the **VBComponents** collection, an enumerated [constant](../../Glossary/vbe-glossary.md#constant) representing a [class module](../../Glossary/vbe-glossary.md#class-module), a form, or a [standard module](../../Glossary/vbe-glossary.md#standard-module).<br/><br/>For the **VBProjects** collection, a standalone project.|

## Remarks

When used on the **LinkedWindows** collection, the **Remove** method removes a window from the collection of currently [linked windows](../../Glossary/vbe-glossary.md#linked-window). The removed window becomes a floating window that has its own [linked window frame](../../Glossary/vbe-glossary.md#linked-window-frame). 

The **Remove** method can only be used on a standalone project. It generates a run-time error if you try to use it on a host project.

> [!IMPORTANT] 
> Objects, properties, and methods for controlling linked windows, linked window frames, and docked windows are included on the Macintosh for compatibility with code written in Windows. However, these language elements generate run-time errors when run on the Macintosh.


## See also

- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]