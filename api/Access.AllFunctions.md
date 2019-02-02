---
title: AllFunctions object (Access)
keywords: vbaac10.chm13250
f1_keywords:
- vbaac10.chm13250
ms.prod: access
api_name:
- Access.AllFunctions
ms.assetid: 1420cf24-906e-7b65-29f3-29a28cdf92cf
ms.date: 02/01/2019
localization_priority: Normal
---


# AllFunctions object (Access)

The **AllFunctions** collection contains an **[AccessObject](Access.AccessObject.md)** object for each function in the **[CurrentData](Access.CurrentData.md)** or **[CodeData](Access.CodeData.md)** object.


## Remarks

The **CurrentData** or **CodeData** object has an **AllFunctions** collection containing **AccessObject** objects that describe instances of all functions specified by the **CurrentData** or **CodeData** objects. For example, you can enumerate the **AllFunctions** collection in Visual Basic to set or return the values of properties of individual **AccessObject** objects in the collection.

You can refer to an individual **AccessObject** object in the **AllFunctions** collection either by referring to the object by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllFunctions** collection, it's better to refer to the function by name because a function's collection index may change.

The **AllFunctions** collection is indexed beginning with zero. If you refer to a function by its index, the first function is AllFunctions(0), the second function is AllFunctions(1), and so on.

To list all open functions in the database, use the **[IsLoaded](Access.AccessObject.IsLoaded.md)** property of each **AccessObject** object in the **AllFunctions** collection. You can then use the **Name** property of each individual **AccessObject** object to return the name of a function.

You can't add or delete an **AccessObject** object from the **AllFunctions** collection.


## Properties

- [Application](Access.AllFunctions.Application.md)
- [Count](Access.AllFunctions.Count.md)
- [Item](Access.AllFunctions.Item.md)
- [Parent](Access.AllFunctions.Parent.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]