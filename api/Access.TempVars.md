---
title: TempVars object (Access)
keywords: vbaac10.chm14073
f1_keywords:
- vbaac10.chm14073
ms.prod: access
api_name:
- Access.TempVars
ms.assetid: aa81b18b-5e9f-ae44-cbcf-55cf6e37b7f6
ms.date: 03/21/2019
localization_priority: Normal
---


# TempVars object (Access)

Represents the collection of **[TempVar](Access.TempVar.md)** objects.


## Remarks

Use the **Add** method or the SetTempVar macro action to create a **TempVar** object.

Use the **Remove** method or the RemoveTempVar macro action to delete a **TempVar** object from the **TempVars** collection.

Use the **RemoveAll** method or the RemoveAllTempVars macro action to delete all **TempVar** objects from the **TempVars** collection.

The **TempVars** collection can store up to 255 **TempVar** objects. If you do not remove a **TempVar** object, it will remain in memory until you close the database. It is a good practice to remove **TempVar** object variables when you are finished using them.

To refer to a **TempVar** object in a collection by its ordinal number or by its **Name** property setting, use the following syntax form:

- **TempVar**![name]
    

## Methods

- [Add](Access.TempVars.Add.md)
- [Remove](Access.TempVars.Remove.md)
- [RemoveAll](Access.TempVars.RemoveAll.md)

## Properties

- [Application](Access.TempVars.Application.md)
- [Count](Access.TempVars.Count.md)
- [Item](Access.TempVars.Item.md)
- [Parent](Access.TempVars.Parent.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
