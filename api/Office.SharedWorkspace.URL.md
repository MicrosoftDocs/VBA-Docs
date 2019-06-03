---
title: SharedWorkspace.URL property (Office)
keywords: vbaof11.chm276011
f1_keywords:
- vbaof11.chm276011
ms.prod: office
api_name:
- Office.SharedWorkspace.URL
ms.assetid: e60e6706-d3f3-1a47-2b8a-82c5d52ddac5
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspace.URL property (Office)

Gets the top-level Uniform Resource Locator (URL) of the shared workspace. Read-only.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**URL**

_expression_ A variable that represents a **[SharedWorkspace](Office.SharedWorkspace.md)** object.


## Return value

String


## Remarks

The URL property returns the address of the shared workspace in this format: `https://server/sites/user/workspace/`. 

The URL property returns a URL-encoded string. For example, a space in the folder name is represented by `%20`. Use a simple function like the following example to replace this escaped character with a space:

`Private Function URLDecode(URLtoDecode As String) As String URLDecode = Replace(URLtoDecode, "%20", " ") End Function`


## Example

The following example displays the base URL of the shared workspace.


```vb
 MsgBox "URL: " & ActiveWorkbook.SharedWorkspaceLink.URL, _ 
        vbInformation + vbOKOnly, "Shared Workspace URL" 

```


## See also

- [SharedWorkspace object members](overview/Library-Reference/sharedworkspace-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]