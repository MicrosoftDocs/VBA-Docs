---
title: Modules object (Access)
keywords: vbaac10.chm12289
f1_keywords:
- vbaac10.chm12289
ms.prod: access
api_name:
- Access.Modules
ms.assetid: f60a9929-4b79-cfed-8fb3-a4869a3afe9f
ms.date: 03/21/2019
localization_priority: Normal
---


# Modules object (Access)

The **Modules** collection contains all open standard modules and class modules in a Microsoft Access database.


## Remarks

All open modules are included in the **Modules** collection, whether they are uncompiled, compiled, or in break mode, or contain the code that's running.

To determine whether an individual **[Module](Access.Module.md)** object represents a standard module or a class module, check the **Module** object's **Type** property.

The **Modules** collection belongs to the Microsoft Access **Application** object.

Individual **Module** objects in the **Modules** collection are indexed beginning with zero.


## Example

The following example illustrates how to use the **Modules** collection to loop through the open modules. The example prints the name of each open module in the Immediate window.

```vb
 
Sub PrintOpenModuleNames() 
 Dim i As Integer 
 Dim modOpenModules As Modules 
 
 Set modOpenModules = Application.Modules 
 
 For i = 0 To modOpenModules.Count - 1 
 
 Debug.Print modOpenModules(i).Name 
 
 Next 
End Sub
```


## Properties

- [Application](Access.Modules.Application.md)
- [Count](Access.Modules.Count.md)
- [Item](Access.Modules.Item.md)
- [Parent](Access.Modules.Parent.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]