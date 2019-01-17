---
title: SharedWorkspaceFile.CreatedBy property (Office)
keywords: vbaof11.chm266002
f1_keywords:
- vbaof11.chm266002
ms.prod: office
api_name:
- Office.SharedWorkspaceFile.CreatedBy
ms.assetid: e16e3e87-7188-7650-db58-d26e7a98d4eb
ms.date: 06/08/2017
localization_priority: Normal
---


# SharedWorkspaceFile.CreatedBy property (Office)

Gets the display name of the member who created the shared workspace object. Read-only.


## Syntax

_expression_. `CreatedBy`

_expression_ A variable that represents a [SharedWorkspaceFile](Office.SharedWorkspaceFile.md) object.


## Return value

String


## Example

The following example lists files in the shared workspace site that were created by users other than the creator of the workspace site.


```vb
 Dim swsFile As Office.SharedWorkspaceFile 
 Dim swsOwner As Office.SharedWorkspaceMember 
 Dim strMemberFiles As String 
 Set swsOwner = ActiveWorkbook.SharedWorkspace.Members(1) 
 For Each swsFile In ActiveWorkbook.SharedWorkspace.Files 
 If swsFile.CreatedBy <> swsOwner.Name Then 
 strMemberFiles = strMemberFiles &amp; swsFile.URL &amp; vbCrLf 
 End If 
 Next 
 MsgBox "These files were created by other users:" &amp; _ 
 vbCrLf &amp; strMemberFiles, _ 
 vbInformation + vbOKOnly, "Files Created by Other Users" 
 Set swsOwner = Nothing 
 Set swsFile = Nothing 

```

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## See also


[SharedWorkspaceFile Object](Office.SharedWorkspaceFile.md)



[SharedWorkspaceFile Object Members](./overview/Library-Reference/sharedworkspacefile-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]