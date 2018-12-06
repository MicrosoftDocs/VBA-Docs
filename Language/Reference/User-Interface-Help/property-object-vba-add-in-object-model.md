---
title: Property object (VBA Add-In Object Model)
keywords: vbob6.chm102045
f1_keywords:
- vbob6.chm102045
ms.prod: office
ms.assetid: 231018ff-4e74-fc67-a69b-0988e5b7517d
ms.date: 12/06/2018
---


# Property object (VBA Add-In Object Model)

Represents the [properties](../../Glossary/vbe-glossary.md#property) of an object that are visible in the [Properties window](../../Glossary/vbe-glossary.md#properties-window) for any given component.

## Remarks

Use the **[Value](value-property-vba-add-in-object-model.md)** property of the **Property** object to return or set the value of a property of a component.

At a minimum, all components have a **[Name](name-property-vba-add-in-object-model.md)** property. The **Value** property returns a [Variant](../../Glossary/vbe-glossary.md#variant-data-type) of the appropriate type. If the value returned is an object, the **Value** property returns the **Properties** collection that contains **Property** objects representing the individual properties of the object. You can access each of the **Property** objects by using the **[Item](item-method-vba-add-in-object-model.md)** method on the returned **Properties** collection.

If the value returned by the **Property** object is an object, you can use the **[Object](object-property-vba-add-in-object-model.md)** property to set the **Property** object to a new object.

## See also

- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)