---
title: SharedWorkspaceLinks.Add method (Office)
keywords: vbaof11.chm271003
f1_keywords:
- vbaof11.chm271003
ms.prod: office
api_name:
- Office.SharedWorkspaceLinks.Add
ms.assetid: 76c1fe99-14de-7276-0c5c-fd54f6d0a6ce
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspaceLinks.Add method (Office)

Adds a link to the list of links in a shared workspace.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**Add** (_URL_, _Description_, _Notes_)

_expression_ Required. A variable that represents a **[SharedWorkspaceLinks](Office.SharedWorkspaceLinks.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _URL_|Required|**String**|The address of the website to which a link is being added.|
| _Description_|Optional|**String**|Description of the link.|
| _Notes_|Optional|**String**|Notes about the link.|

## Return value

SharedWorkspaceLink


## Example

The following example adds a new link to the links collection of the shared workspace.


```vb
    Dim swsLink As Office.SharedWorkspaceLink 
    Set swsLink = ActiveWorkbook.SharedWorkspace.Links.Add( _ 
        "https://msdn.microsoft.com", _ 
        "Microsoft Developer Network Home Page", _ 
        "My favorite developer site!") 
    MsgBox "New link: " & swsLink.Description, _ 
        vbInformation + vbOKOnly, _ 
        "New Link in Shared Workspace" 
    Set swsLink = Nothing 

```


## See also

- [SharedWorkspaceLinks object members](overview/Library-Reference/sharedworkspacelinks-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]