---
title: SharedWorkspaceMember.Email property (Office)
keywords: vbaof11.chm272003
f1_keywords:
- vbaof11.chm272003
ms.prod: office
api_name:
- Office.SharedWorkspaceMember.Email
ms.assetid: 3539becc-bde4-9331-432c-e907523975a7
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspaceMember.Email property (Office)

Gets the email name of the specified **SharedWorkspaceMember** in the format user@domain.com. Read-only.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**Email**

_expression_ An expression that returns a **[SharedWorkspaceMember](Office.SharedWorkspaceMember.md)** object.


## Example

The following example extracts the email domain name from the **Email** property of each shared workspace member and lists members who have email addresses at the `"example.com"` domain.


```vb
Dim swsMember As Office.SharedWorkspaceMember 
    Dim strEmailDomain As String 
    Dim strMemberList As String 
    For Each swsMember In ActiveWorkbook.SharedWorkspace.Members 
        strEmailDomain = LCase(Right(swsMember.Email, _ 
            Len(swsMember.Email) - InStr(swsMember.Email, "@"))) 
        If strEmailDomain = "example.com" Then 
            strMemberList = strMemberList & swsMember.Email & vbCrLf 
        End If 
    Next 
    MsgBox strMemberList, vbInformation + vbOKOnly, _ 
        "Members with example.com email" 
    Set swsMember = Nothing
```


## See also

- [SharedWorkspaceMember object members](overview/Library-Reference/sharedworkspacemember-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]