---
title: SharedWorkspaceMembers object (Office)
keywords: vbaof11.chm273000
f1_keywords:
- vbaof11.chm273000
ms.prod: office
api_name:
- Office.SharedWorkspaceMembers
ms.assetid: 2d0e6ce0-79ef-3030-b1af-465428314b15
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspaceMembers object (Office)

A collection of the **[SharedWorkspaceMember](Office.SharedWorkspaceMember.md)** objects in the current shared workspace site.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Example

Use the **[Members](Office.SharedWorkspace.Members.md)** property of the **[SharedWorkspace](Office.SharedWorkspace.md)** object to return a **SharedWorkspaceMembers** collection.


```vb
    Dim swsMembers As Office.SharedWorkspaceMembers 
    Set swsMembers = ActiveWorkbook.SharedWorkspace.Members 
    MsgBox "There are " & swsMembers.Count & _ 
        " member(s) in the current shared workspace.", _ 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsMembers = Nothing 

```


## See also

- [SharedWorkspaceMembers object members](overview/Library-Reference/sharedworkspacemembers-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]