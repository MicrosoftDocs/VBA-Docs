---
title: Reference object (VBA Add-In Object Model)
keywords: vbob6.chm104053
f1_keywords:
- vbob6.chm104053
ms.prod: office
ms.assetid: 559d4da0-624f-a574-575d-768155c89c72
ms.date: 12/06/2018
---


# Reference object (VBA Add-In Object Model)

Represents a reference to a [type library](../../Glossary/vbe-glossary.md#type-library) or a [project](../../Glossary/vbe-glossary.md#project).

## Remarks

Use the **Reference** object to verify whether a reference is still valid.

The **[IsBroken](isbroken-property-vba-add-in-object-model.md)** property returns **True** if the reference no longer points to a valid reference. 

The **[BuiltIn](builtin-property-vba-add-in-object-model.md)** property returns **True** if the reference is a default reference that can't be moved or removed. 

Use the **[Name](name-property-vba-add-in-object-model.md)** property to determine if the reference you want to add or remove is the correct one.

## See also

- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)