---
title: SharedWorkspaceLink.URL property (Office)
keywords: vbaof11.chm270001
f1_keywords:
- vbaof11.chm270001
ms.prod: office
api_name:
- Office.SharedWorkspaceLink.URL
ms.assetid: 92104c43-43b8-5f59-e0c0-91313d8f5e35
ms.date: 06/08/2017
localization_priority: Normal
---


# SharedWorkspaceLink.URL property (Office)

Gets the top-level Uniform Resource Locator (URL) of the shared workspace link. Read/write.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_. `URL`

_expression_ A variable that represents a [SharedWorkspaceLink](Office.SharedWorkspaceLink.md) object.


## Return value

String


## Remarks

The URL property returns the address of the shared workspace in this format:  `https://server/sites/user/workspace/`. The URL property returns a URL-encoded string. For example, a space in the folder name is represented by %20. Use a simple function like the following example to replace this escaped character with a space. `Private Function URLDecode(URLtoDecode As String) As String URLDecode = Replace(URLtoDecode, "%20", " ") End Function`


## Example

The following example displays the URL of the link to the shared workspace.


```vb
MsgBox "URL: " &amp; ActiveWorkbook.SharedWorkspaceLink.URL, _ 
        vbInformation + vbOKOnly, "Shared Workspace Link URL"
```


## See also


[SharedWorkspaceLink Object](Office.SharedWorkspaceLink.md)



[SharedWorkspaceLink Object Members](./overview/Library-Reference/sharedworkspacelink-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]