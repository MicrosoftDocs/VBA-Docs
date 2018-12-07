---
title: VBComponents property
keywords: vbob6.chm102023
f1_keywords:
- vbob6.chm102023
ms.prod: office
api_name:
- Office.VBComponents
ms.assetid: cf5dea3b-e583-8547-9448-0fd23e5623ee
ms.date: 12/06/2018
---


# VBComponents property

Returns a collection of the components contained in a project.

### Remarks

Use the **[VBComponents](vbcomponents-collection-vba-add-in-object-model.md)** collection to access, add, or remove components in a project. A component can be a [form](../../Glossary/vbe-glossary.md#form), [module](../../Glossary/vbe-glossary.md#module), or [class](../../Glossary/vbe-glossary.md#class). The **VBComponents** collection is a standard [collection](../../Glossary/vbe-glossary.md#collection) that can be used in a **For… Each** block.

You can use the **[Parent](parent-property-vba-add-in-object-model.md)** property to return the project that the **VBComponents** collection is in.

In Visual Basic for Applications, you can use the **[Import](import-method-vba-add-in-object-model.md)** method to add a component to a project from a file.

### See also

- [VBComponent object](vbcomponent-object-vba-add-in-object-model.md)
- [SelectedVBComponent property](selectedvbcomponent-property-vba-add-in-object-model.md)
- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)