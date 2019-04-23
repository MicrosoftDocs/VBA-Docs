---
title: DeleteLines method (VBA Add-In Object Model)
keywords: vbob6.chm104014
f1_keywords:
- vbob6.chm104014
ms.prod: office
ms.assetid: b6e1bd5d-23b2-0bc4-bcc6-b7e371df4b93
ms.date: 12/06/2018
localization_priority: Normal
---


# DeleteLines method (VBA Add-In Object Model)

Deletes a single line or a specified range of lines.

## Syntax

_object_.**DeleteLines** (_startline_, [ _count_ ])

<br/>

The **DeleteLines** syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.|
| _startline_|Required. A [Long](../../Glossary/vbe-glossary.md#long-data-type) specifying the first line you want to delete.|
| _count_|Optional. A **Long** specifying the number of lines you want to delete.|

## Remarks

If you don't specify how many lines you want to delete, **DeleteLines** deletes one line.

## See also

- [CodeModule object](../visual-basic-add-in-model/objects-visual-basic-add-in-model.md#codemodule)
- [Module.DeleteLines method (Access)](../../../api/access.module.deletelines.md)
- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]