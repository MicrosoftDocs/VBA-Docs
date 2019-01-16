---
title: SharedWorkspaceFile.ModifiedBy property (Office)
keywords: vbaof11.chm266004
f1_keywords:
- vbaof11.chm266004
ms.prod: office
api_name:
- Office.SharedWorkspaceFile.ModifiedBy
ms.assetid: d6533854-ddd9-3a41-b74b-94f282779236
ms.date: 06/08/2017
localization_priority: Normal
---


# SharedWorkspaceFile.ModifiedBy property (Office)

Gets the name of the user who last modified the object. Read-only.


## Syntax

_expression_. `ModifiedBy`

_expression_ A variable that represents a [SharedWorkspaceFile](Office.SharedWorkspaceFile.md) object.


## Return value

String


## Remarks

For shared workspace objects, the  **ModifiedBy** property returns the display name stored in the **Name** property of the **SharedWorkspaceMember** object.


## Example

The following example lists the files in a shared workspace site that were last modified by users other than the creator of the workspace site.


```vb
 Dim swsFile As Office.SharedWorkspaceFile 
 Dim swsOwner As Office.SharedWorkspaceMember 
 Dim strMemberFiles As String 
 Set swsOwner = ActiveWorkbook.SharedWorkspace.Members(1) 
 For Each swsFile In ActiveWorkbook.SharedWorkspace.Files 
 If swsFile.ModifiedBy <> swsOwner.Name Then 
 strMemberFiles = strMemberFiles &amp; swsFile.URL &amp; vbCrLf 
 End If 
 Next 
 MsgBox "These files were last modified by other users:" &amp; _ 
 vbCrLf &amp; strMemberFiles, _ 
 vbInformation + vbOKOnly, "Files Modified by Other Users" 
 Set swsOwner = Nothing 
 Set swsFile = Nothing 

```

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## See also


[SharedWorkspaceFile Object](Office.SharedWorkspaceFile.md)



[SharedWorkspaceFile Object Members](./overview/Library-Reference/sharedworkspacefile-members-office.md)

