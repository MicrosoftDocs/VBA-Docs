---
title: SaveAs method (VBA Add-In Object Model)
keywords: vbob6.chm102017
f1_keywords:
- vbob6.chm102017
ms.prod: office
ms.assetid: 622aa652-8093-be64-4128-9ad2c7fd1fe8
ms.date: 12/06/2018
localization_priority: Normal
---


# SaveAs method (VBA Add-In Object Model)

Saves a project to a given location by using a new filename.

## Syntax

_object_.**SaveAs** (_newfilename_) **As String**

<br/>

The **SaveAs** method syntax has these parts.

|Part|Description|
|:-----|:-----|
| _object_|An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.|
| _newfilename_|Required. A [string expression](../../Glossary/vbe-glossary.md#string-expression) specifying the new filename for the component to be saved.|

## Remarks

If a new path name is given, it is used. Otherwise, the old path name is used. If the new filename is invalid or refers to a read-only file, an error occurs.

The **SaveAs** method can only be used on standalone projects. It generates a run-time error if you use it with a host project.

## See also

- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]