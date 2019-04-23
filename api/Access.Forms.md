---
title: Forms object (Access)
keywords: vbaac10.chm12355
f1_keywords:
- vbaac10.chm12355
ms.prod: access
api_name:
- Access.Forms
ms.assetid: a41af7be-873c-ef8b-20cd-24b78a25b5ca
ms.date: 03/20/2019
localization_priority: Normal
---


# Forms object (Access)

The **Forms** collection contains all the currently open forms in a Microsoft Access database.


## Remarks

Use the **Forms** collection in Visual Basic or in an expression to refer to forms that are currently open. For example, you can enumerate the **Forms** collection to set or return the values of properties of individual forms in the collection.

You can refer to an individual **[Form](access.form.md)** object in the **Forms** collection either by referring to the form by name, or by referring to its index within the collection. If you want to refer to a specific form in the **Forms** collection, it's better to refer to the form by name because a form's collection index may change.

The **Forms** collection is indexed beginning with zero. If you refer to a form by its index, the first form opened is Forms(0), the second form opened is Forms(1), and so on. If you opened Form1 and then opened Form2, Form2 would be referenced in the **Forms** collection by its index as Forms(1). If you then closed Form1, Form2 would be referenced in the **Forms** collection by its index as Forms(0).

> [!NOTE] 
> To list all forms in the database, whether open or closed, enumerate the **AllForms** collection of the **[CurrentProject](Access.CurrentProject.md)** object. You can then use the **Name** property of each individual **[AccessObject](Access.AccessObject.md)** object to return the name of a form.

You can't add or delete a **Form** object from the **Forms** collection.


## Properties

- [Application](Access.Forms.Application.md)
- [Count](Access.Forms.Count.md)
- [Item](Access.Forms.Item.md)
- [Parent](Access.Forms.Parent.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
