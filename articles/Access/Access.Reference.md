---
title: Reference Object (Access)
keywords: vbaac10.chm12628
f1_keywords:
- vbaac10.chm12628
ms.prod: access
api_name:
- Access.Reference
ms.assetid: 87853230-294e-7ab8-4aae-78b094b5e584
ms.date: 06/08/2017
---


# Reference Object (Access)

The  **Reference** object refers to a reference set to another application's or project's type library.


## Remarks

When you create a  **Reference** object, you set a reference dynamically from Visual Basic.

The  **Reference** object is a member of the **References** collection. To refer to a particular **Reference** object in the **References** collection, use any of the following syntax forms.



|**Syntax**|**Description**|
|:-----|:-----|
|**References** ! _referencename_|The  _referencename_ argument is the name of the **Reference** object.|
|**References** (" _referencename_")|The  _referencename_ argument is the name of the **Reference** object.|
|**References** ( _index_)|The  _index_ argument is the object's numerical position within the collection.|

 **Note**  The following example refers to the  **Reference** object that represents the reference to the Microsoft Access type library:




```
Dim ref As Reference 
Set ref = References!Access
```


## Properties



|**Name**|
|:-----|
|[BuiltIn](Access.Reference.BuiltIn.md)|
|[Collection](Access.Reference.Collection.md)|
|[FullPath](Access.Reference.FullPath.md)|
|[Guid](Access.Reference.Guid.md)|
|[IsBroken](Access.Reference.IsBroken.md)|
|[Kind](Access.Reference.Kind.md)|
|[Major](Access.Reference.Major.md)|
|[Minor](Access.Reference.Minor.md)|
|[Name](Access.Reference.Name.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
