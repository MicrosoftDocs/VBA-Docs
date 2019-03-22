---
title: Reference.GUID property (Access)
keywords: vbaac10.chm12631
f1_keywords:
- vbaac10.chm12631
ms.prod: access
api_name:
- Access.Reference.Guid
ms.assetid: a5419b60-f113-2c56-ff74-62c9ff8cc868
ms.date: 03/23/2019
localization_priority: Normal
---


# Reference.GUID property (Access)

The **GUID** property of a **Reference** object returns a GUID that identifies a type library in the Windows Registry. Read-only **String**.


## Syntax

_expression_.**GUID**

_expression_ A variable that represents a **[Reference](Access.Reference.md)** object.


## Remarks

Every type library has an associated GUID that is stored in the Registry. When you set a reference to a type library, Microsoft Access uses the type library's GUID to identify the type library.

You can use the **[AddFromGUID](Access.References.AddFromGuid.md)** method to create a **Reference** object from a type library's GUID.


## Example

The following example prints the value of the **FullPath**, **GUID**, **IsBroken**, **Major**, and **Minor** properties for each **Reference** object in the **[References](Access.References.md)** collection.

```vb
Sub ReferenceProperties() 
 Dim ref As Reference 
 
 ' Enumerate through References collection. 
 For Each ref In References 
 ' Check IsBroken property. 
 If ref.IsBroken = False Then 
 Debug.Print "Name: ", ref.Name 
 Debug.Print "FullPath: ", ref.FullPath 
 Debug.Print "Version: ", ref.Major & "." & ref.Minor 
 Else 
 Debug.Print "GUIDs of broken references:" 
 Debug.Print ref.GUID 
 EndIf 
 Next ref 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]