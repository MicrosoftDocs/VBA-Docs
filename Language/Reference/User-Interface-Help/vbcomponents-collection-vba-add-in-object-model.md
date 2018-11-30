---
title: VBComponents collection (VBA Add-In Object Model)
keywords: vbob6.chm1070945
f1_keywords:
- vbob6.chm1070945
ms.prod: office
ms.assetid: de087f44-949a-949a-9703-244ea076480e
ms.date: 11/29/2018
---


# VBComponents collection (VBA Add-In Object Model)

Represents the components contained in a [project](../../Glossary/vbe-glossary.md#project).

## Remarks

Use the **VBComponents** collection to access, add, or remove components in a project. A component can be a [form](../../Glossary/vbe-glossary.md#form), [module](../../Glossary/vbe-glossary.md#module), or [class](../../Glossary/vbe-glossary.md#class). The **VBComponents** collection is a standard [collection](../../Glossary/vbe-glossary.md#collection) that can be used in a **For Each** block.

You can use the **[Parent](parent-property-vba-add-in-object-model.md)** property to return the project the **VBComponents** collection is in.

In Visual Basic for Applications, you can use the **[Import](import-method-vba-add-in-object-model.md)** method to add a component to a project from a file.

## See also

- [VBComponent object](vbcomponent-object-vba-add-in-object-model.md)
- [VBComponents property](vbcomponents-property.md)
- [SelectedVBComponent property](selectedvbcomponent-property-vba-add-in-object-model.md)