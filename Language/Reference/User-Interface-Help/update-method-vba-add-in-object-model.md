---
title: Update method (VBA Add-In Object Model)
keywords: vbob6.chm102251
f1_keywords:
- vbob6.chm102251
ms.prod: office
ms.assetid: c88ee513-6d8e-9c40-2999-4cc217fc3fc8
ms.date: 12/06/2018
localization_priority: Normal
---


# Update method (VBA Add-In Object Model)

Refreshes the contents of the **AddIns** collection from the add-ins listed in the Vbaddin.ini file in the same manner as if the user had opened the **[Add-In Manager](add-in-manager-dialog-box.md)** dialog box.

## Syntax

_object_.**Update**

The _object_ placeholder represents an [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.

## Remarks

All add-ins listed in the Vbaddin.ini file must be registered ActiveX components in the Registry before they can be used in Visual Basic.

## See also

- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]