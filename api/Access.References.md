---
title: References object (Access)
keywords: vbaac10.chm12648
f1_keywords:
- vbaac10.chm12648
ms.prod: access
api_name:
- Access.References
ms.assetid: ac020382-4ece-f138-d1b9-d05b0fe0f523
ms.date: 03/21/2019
localization_priority: Normal
---


# References object (Access)

The **References** collection contains **[Reference](access.reference.md)** objects representing each reference that's currently set.


## Remarks

The **Reference** objects in the **References** collection correspond to the list of references in the **References** dialog box, available by choosing **References** on the **Tools** menu. Each **Reference** object represents one selected reference in the list. References that appear in the **References** dialog box but haven't been selected aren't in the **References** collection.

You can enumerate through the **References** collection by using the **For Each...Next** statement.

The **References** collection belongs to the Microsoft Access **Application** object.

Individual **Reference** objects in the **References** collection are indexed beginning with 1.

## Events

- [ItemAdded](Access.References.ItemAdded.md)
- [ItemRemoved](Access.References.ItemRemoved.md)

## Methods

- [AddFromFile](Access.References.AddFromFile.md)
- [AddFromGuid](Access.References.AddFromGuid.md)
- [Item](Access.References.Item.md)
- [Remove](Access.References.Remove.md)

## Properties

- [Count](Access.References.Count.md)
- [Parent](Access.References.Parent.md)


## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]