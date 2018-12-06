---
title: CodeModule property (VBA Add-In Object Model)
keywords: vbob6.chm1071208
f1_keywords:
- vbob6.chm1071208
ms.prod: office
ms.assetid: 5e99b614-7207-f577-49dd-5199cb4d9373
ms.date: 12/06/2018
---


# CodeModule property (VBA Add-In Object Model)

Returns an object representing the code behind the component. Read-only.

## Remarks

The **CodeModule** property returns **[Nothing](nothing-keyword.md)** if the component doesn't have a [code module](../../Glossary/vbe-glossary.md#code-module) associated with it.

> [!NOTE] 
> The **[CodePane](codepane-object-vba-add-in-object-model.md)** object represents a visible code window. A given component can have several **CodePane** objects. 
> 
> The **[CodeModule](codemodule-object-vba-add-in-object-model.md)** object represents the code within a component. A component can only have one **CodeModule** object.

## See also

- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)