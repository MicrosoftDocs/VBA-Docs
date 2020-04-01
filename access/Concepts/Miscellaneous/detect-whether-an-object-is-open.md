---
title: Detect whether an object is open
ms.prod: access
ms.assetid: 9378430d-752b-1ede-96af-235c7e79a06f
ROBOTS: INDEX
ms.date: 06/08/2019
localization_priority: Normal
---


# Detect whether an object is open

It is often necessary to know whether a particular database object is open before you can edit the object programmatically. The following example illustrates how to use the **[SysCmd](../../../api/Access.Application.SysCmd.md)** method with the **acSysCmdGetObjectState** action to determine whether a database object is open.

The example function, **IsObjectLoaded**, accepts two parameters. The _strObjectName_ parameter is the name of the database object to check for. The _strObjectType_ parameter is an **[AcObjectType](../../../api/Access.AcObjectType.md)** constant that specifies the type of database object to check for. The **IsObjectLoaded** function returns **True** if the specified database object is open, and returns **False** if it is not open.

```vb
 
Function IsObjectLoaded(ByVal strObjectName As String, ByVal strObjectType As AcObjectType) As Boolean 
     
    If SysCmd(acSysCmdGetObjectState, strObjectType, strObjectName) <> 0 Then 
         
       ' The object is open. 
        IsObjectLoaded = True 
    Else 
 
       ' The object is not open. 
        IsObjectLoaded = False 
    End If 
     
End Function
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]