---
title: Close method (VBA Add-In Object Model)
keywords: vbob6.chm100111
f1_keywords:
- vbob6.chm100111
ms.prod: office
ms.assetid: e3c951ed-032b-9e4b-ba1b-a802f42d3544
ms.date: 12/06/2018
localization_priority: Normal
---


# Close method (VBA Add-In Object Model)

Closes and destroys a [window](../visual-basic-add-in-model/objects-visual-basic-add-in-model.md#window).

## Syntax

_object_.**Close**

The _object_ placeholder is an [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.

## Remarks

The following types of windows respond to the **Close** method in different ways:

- For a window that is a [code pane](../../Glossary/vbe-glossary.md#code-pane), **Close** destroys the code pane.
    
- For a window that is a [designer](../../Glossary/vbe-glossary.md#designer), **Close** destroys the contained designer.
    
- For windows that are always available on the **[View](view-menu.md)** menu, **Close** hides the window.
    
## See also

- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]