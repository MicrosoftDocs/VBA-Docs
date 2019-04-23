---
title: Export method (VBA Add-In Object Model)
keywords: vbob6.chm102194
f1_keywords:
- vbob6.chm102194
ms.prod: office
ms.assetid: 46cab37a-4390-219c-68f8-05cbb59c0450
ms.date: 12/06/2018
localization_priority: Normal
---


# Export method (VBA Add-In Object Model)

Saves a component as a separate file or files.

## Syntax

_object_.**Export** (_filename_)

<br/>

The **Export** syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.|
| _filename_|Required. A [String](../../Glossary/vbe-glossary.md#string-data-type) specifying the name of the file that you want to export the component to.|

## Remarks

When you use the **Export** method to save a component as a separate file or files, use a file name that doesn't already exist; otherwise, an error occurs.

## See also

- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]