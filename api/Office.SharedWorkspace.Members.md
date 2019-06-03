---
title: SharedWorkspace.Members property (Office)
keywords: vbaof11.chm276002
f1_keywords:
- vbaof11.chm276002
ms.prod: office
api_name:
- Office.SharedWorkspace.Members
ms.assetid: a53cfd41-36ca-73e4-08b2-306569f26979
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspace.Members property (Office)

Gets a **[SharedWorkspaceMembers](Office.SharedWorkspaceMembers.md)** collection that represents the list of members in the current shared workspace. Read-only.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**Members**

_expression_ A variable that represents a **[SharedWorkspace](Office.SharedWorkspace.md)** object.


## Example

The following example lists the members in the current shared workspace.


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

- [SharedWorkspace object members](overview/Library-Reference/sharedworkspace-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]