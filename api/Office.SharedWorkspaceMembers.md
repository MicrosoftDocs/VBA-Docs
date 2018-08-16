---
title: SharedWorkspaceMembers Object (Office)
keywords: vbaof11.chm273000
f1_keywords:
- vbaof11.chm273000
ms.prod: office
api_name:
- Office.SharedWorkspaceMembers
ms.assetid: 2d0e6ce0-79ef-3030-b1af-465428314b15
ms.date: 06/08/2017
---


# SharedWorkspaceMembers Object (Office)

A collection of the  **[SharedWorkspaceMember](Office.SharedWorkspaceMember.md)** objects in the current shared workspace site.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Example

Use the  **[Members](Office.SharedWorkspace.Members.md)** property of the **[SharedWorkspace](Office.SharedWorkspace.md)** object to return a **SharedWorkspaceMembers** collection.


```vb
    Dim swsMembers As Office.SharedWorkspaceMembers 
    Set swsMembers = ActiveWorkbook.SharedWorkspace.Members 
    MsgBox "There are " &amp; swsMembers.Count &amp; _ 
        " member(s) in the current shared workspace.", _ 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsMembers = Nothing 

```


## Methods



|**Name**|
|:-----|
|[Add](Office.SharedWorkspaceMembers.Add.md)|

## Properties



|**Name**|
|:-----|
|[Application](Office.SharedWorkspaceMembers.Application.md)|
|[Count](Office.SharedWorkspaceMembers.Count.md)|
|[Creator](Office.SharedWorkspaceMembers.Creator.md)|
|[Item](Office.SharedWorkspaceMembers.Item.md)|
|[ItemCountExceeded](Office.SharedWorkspaceMembers.ItemCountExceeded.md)|
|[Parent](Office.SharedWorkspaceMembers.Parent.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)
