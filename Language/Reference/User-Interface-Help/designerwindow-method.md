---
title: DesignerWindow method (VBA)
keywords: vbob6.chm102199
f1_keywords:
- vbob6.chm102199
ms.prod: office
api_name:
- Office.DesignerWindow
ms.assetid: 1a116dab-56ce-087e-1789-614a3709c9cc
ms.date: 12/06/2018
localization_priority: Normal
---


# DesignerWindow method

Returns the **[Window](window-object-vba-add-in-object-model.md)** object that represents the component's [designer](../../Glossary/vbe-glossary.md#designer).

## Syntax

_object_.**DesignerWindow**

The _object_ placeholder is an [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.

## Remarks

If the component supports a designer but doesn't have an open designer, using the **DesignerWindow** method creates the designer, but it isn't visible. To make the window visible, set the **Window** object's **[Visible](visible-property-vba-add-in-object-model.md)** property to **True**.

## See also

- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]