---
title: SharedWorkspace.RemoveDocument method (Office)
keywords: vbaof11.chm276015
f1_keywords:
- vbaof11.chm276015
ms.prod: office
api_name:
- Office.SharedWorkspace.RemoveDocument
ms.assetid: 4bfb27d7-6fdd-9350-70d2-9c60d75020eb
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspace.RemoveDocument method (Office)

Removes the active document from the shared workspace site.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**RemoveDocument**

_expression_ A variable that represents a **[SharedWorkspace](Office.SharedWorkspace.md)** object.


## Remarks

If the user does not have permission to remove the shared workspace document from the server, the server copy remains intact, but the local copy of the document is disconnected from the shared workspace. In the case where the document was opened directly from the workspace and then removed from the workspace by using **RemoveDocument**, the document must be saved to another location before closing; otherwise, it remains in the workspace.

Use the **[Disconnect](Office.SharedWorkspace.Disconnect.md)** method to detach the local copy of the document from the shared workspace without removing the shared copy.


## Example

The following example determines whether the active document is connected to a shared workspace, and then offers the user the option of removing the document from the workspace site.


```vb
    Dim r As Long 
    If ActiveWorkbook.SharedWorkspace.Connected Then 
        r = MsgBox("Are you sure you want to remove this document?", _ 
            vbQuestion + vbOKCancel, "Are you sure?") 
        If r = vbOK Then 
            ActiveWorkbook.SharedWorkspace.RemoveDocument 
            MsgBox "The document is removed.", _ 
                vbInformation + vbOKOnly, "Removed" 
        Else 
            MsgBox "Removal canceled.", _ 
                vbInformation + vbOKOnly, "Still In Workspace" 
        End If 
    Else 
        MsgBox "The active document is not connected to a shared workspace.", _ 
            vbInformation + vbOKOnly, "Not Connected" 
    End If 

```


## See also

- [SharedWorkspace object members](overview/Library-Reference/sharedworkspace-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]