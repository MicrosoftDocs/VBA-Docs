---
title: Import method (VBA Add-In Object Model)
keywords: vbob6.chm1098974
f1_keywords:
- vbob6.chm1098974
ms.prod: office
ms.assetid: 7ca2c050-6403-bd58-03a9-05111390d398
ms.date: 12/06/2018
localization_priority: Normal
---


# Import method (VBA Add-In Object Model)

Adds a component to a [project](../../Glossary/vbe-glossary.md#project) from a file; returns the newly added component.

## Syntax

_object_.**Import** (_filename_) **As VBComponent**

<br/>

The **Import** syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.|
| _filename_|Required. A [String](../../Glossary/vbe-glossary.md#string-data-type) specifying the path and file name of the component that you want to import the component from.|

## Remarks

You can use the **Import** method to add a component, [form](../../Glossary/vbe-glossary.md#form), [module](../../Glossary/vbe-glossary.md#module), [class](../../Glossary/vbe-glossary.md#class), and so on, to your project.

## See also

- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]