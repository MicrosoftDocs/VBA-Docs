---
title: Reference object (Access)
keywords: vbaac10.chm12628
f1_keywords:
- vbaac10.chm12628
ms.prod: access
api_name:
- Access.Reference
ms.assetid: 87853230-294e-7ab8-4aae-78b094b5e584
ms.date: 03/21/2019
localization_priority: Normal
---


# Reference object (Access)

The **Reference** object refers to a reference set to another application's or project's type library.


## Remarks

When you create a **Reference** object, you set a reference dynamically from Visual Basic.

The **Reference** object is a member of the **References** collection. To refer to a particular **Reference** object in the **References** collection, use any of the following syntax forms.

|Syntax|Description|
|:-----|:-----|
|**References**!_referencename_|The _referencename_ argument is the name of the **Reference** object.|
|**References**("_referencename_")|The _referencename_ argument is the name of the **Reference** object.|
|**References**(_index_)|The _index_ argument is the object's numerical position within the collection.|

The following example refers to the **Reference** object that represents the reference to the Microsoft Access type library.

```vb
Dim ref As Reference 
Set ref = References!Access
```


## Properties

- [BuiltIn](Access.Reference.BuiltIn.md)
- [Collection](Access.Reference.Collection.md)
- [FullPath](Access.Reference.FullPath.md)
- [Guid](Access.Reference.Guid.md)
- [IsBroken](Access.Reference.IsBroken.md)
- [Kind](Access.Reference.Kind.md)
- [Major](Access.Reference.Major.md)
- [Minor](Access.Reference.Minor.md)
- [Name](Access.Reference.Name.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]