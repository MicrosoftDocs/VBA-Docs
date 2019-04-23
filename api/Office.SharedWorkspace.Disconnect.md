---
title: SharedWorkspace.Disconnect method (Office)
keywords: vbaof11.chm276016
f1_keywords:
- vbaof11.chm276016
ms.prod: office
api_name:
- Office.SharedWorkspace.Disconnect
ms.assetid: a742bdc5-4fe1-fa51-bdb9-290fd7179ea7
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspace.Disconnect method (Office)

Disconnects the local copy of the active document from the shared workspace site.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**Disconnect**

_expression_ A variable that represents a **[SharedWorkspace](Office.SharedWorkspace.md)** object.


## Remarks

Use the **Disconnect** method to detach the local copy of the active document from the shared workspace site. This method leaves the shared document on the server; however, the local copy will no longer be synchronized with the shared copy and will no longer benefit from the other collaboration features of the shared workspace. Use the **RemoveDocument** method to remove the shared document from the server.


## Example

The following example determines whether the active document is connected to a shared workspace site, and then offers the user the option of disconnecting it.


```vb
   Dim r As Long 
    If ActiveWorkbook.SharedWorkspace.Connected Then 
        r = MsgBox("Are you sure you want to disconnect this document?", _ 
            vbQuestion + vbOKCancel, "Are you sure?") 
        If r = vbOK Then 
            ActiveWorkbook.SharedWorkspace.Disconnect 
            MsgBox "The document has been disconnected.", _ 
                vbInformation + vbOKOnly, "Disconnected" 
        Else 
            MsgBox "Disconnect canceled.", _ 
                vbInformation + vbOKOnly, "Still Connected" 
        End If 
    Else 
        MsgBox "The active document is not connected to a shared workspace.", _ 
            vbInformation + vbOKOnly, "Not Connected" 
    End If 

```


## See also

- [SharedWorkspace object members](overview/Library-Reference/sharedworkspace-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]