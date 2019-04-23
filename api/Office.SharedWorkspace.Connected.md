---
title: SharedWorkspace.Connected property (Office)
keywords: vbaof11.chm276012
f1_keywords:
- vbaof11.chm276012
ms.prod: office
api_name:
- Office.SharedWorkspace.Connected
ms.assetid: 071502b9-c4f7-45f5-062b-818d5859708e
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspace.Connected property (Office)

Gets a **Boolean** value that indicates whether or not the active document is currently saved in and connected to a shared workspace. Read-only.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**Connected**

_expression_ A variable that represents a **[SharedWorkspace](Office.SharedWorkspace.md)** object.


## Remarks

Use the **[Disconnect](Office.SharedWorkspace.Disconnect.md)** method of the **SharedWorkspace** object to disconnect the local copy of the active document from the shared workspace. Use the **[RemoveDocument](Office.SharedWorkspace.RemoveDocument.md)** method to remove the document from the shared workspace.


## Example

The following example checks the **Connected** property to determine whether the active document is already saved in a shared workspace.


```vb
    If ActiveWorkbook.SharedWorkspace.Connected Then 
        MsgBox "This document is already saved in a shared workspace." 
    End If 

```


## See also

- [SharedWorkspace object members](overview/Library-Reference/sharedworkspace-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]