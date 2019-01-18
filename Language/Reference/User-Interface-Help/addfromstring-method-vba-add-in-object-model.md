---
title: AddFromString method (VBA Add-In Object Model)
keywords: vbob6.chm1098959
f1_keywords:
- vbob6.chm1098959
ms.prod: office
ms.assetid: a3ad95b2-6327-ba69-71d5-17d4f693462c
ms.date: 12/06/2018
localization_priority: Normal
---


# AddFromString method (VBA Add-In Object Model)

Adds text to a [module](../../Glossary/vbe-glossary.md#module).

## Syntax

_object_.**AddFromString**

The _object_ placeholder is an [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.

## Remarks

The **AddFromString** method inserts the text starting on the line preceding the first [procedure](../../Glossary/vbe-glossary.md#procedure) in the module. If the module doesn't contain procedures, **AddFromString** places the text at the end of the module.

## See also

- [CodeModule object](../visual-basic-add-in-model/objects-visual-basic-add-in-model.md#codemodule)
- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]