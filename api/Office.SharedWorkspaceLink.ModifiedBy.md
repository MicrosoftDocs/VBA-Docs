---
title: SharedWorkspaceLink.ModifiedBy property (Office)
keywords: vbaof11.chm270006
f1_keywords:
- vbaof11.chm270006
ms.prod: office
api_name:
- Office.SharedWorkspaceLink.ModifiedBy
ms.assetid: 3070460c-c3af-ff17-19b7-25a3c6339628
ms.date: 06/08/2017
localization_priority: Normal
---


# SharedWorkspaceLink.ModifiedBy property (Office)

Gets the name of the user who last modified the object. Read-only.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_. `ModifiedBy`

_expression_ A variable that represents a [SharedWorkspaceLink](Office.SharedWorkspaceLink.md) object.


## Return value

String


## Remarks

For shared workspace objects, the  **ModifiedBy** property returns the display name stored in the **Name** property of the **SharedWorkspaceMember** object.


## Example

The following example lists the links in a shared workspace site that were modified by a particular users.


```vb
    Dim swsLink As Office.SharedWorkspaceLink 
    Dim swsOwner As Office.SharedWorkspaceMember 
    Dim strMemberFiles As String 
    Dim strUser As String 
    strUser = "Nancy Davolio" 
    Set swsOwner = ActiveWorkbook.SharedWorkspace.Members(1) 
    For Each swsLink In ActiveWorkbook.SharedWorkspace.Links 
        If swsLink.ModifiedBy = strUser Then 
            strMemberlinks = strMemberlinks &amp; swsLink.URL &amp; vbCrLf 
        End If 
    Next 
    MsgBox "These links were modified by " &amp; _ 
        strUser &amp; vbCrLf &amp; strMemberlinks, _ 
        vbInformation + vbOKOnly, "Modified Links" 
    Set swsOwner = Nothing 
    Set swsLink = Nothing 

```


## See also


[SharedWorkspaceLink Object](Office.SharedWorkspaceLink.md)



[SharedWorkspaceLink Object Members](./overview/Library-Reference/sharedworkspacelink-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]