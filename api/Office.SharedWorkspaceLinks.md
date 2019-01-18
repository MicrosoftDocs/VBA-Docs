---
title: SharedWorkspaceLinks object (Office)
keywords: vbaof11.chm271000
f1_keywords:
- vbaof11.chm271000
ms.prod: office
api_name:
- Office.SharedWorkspaceLinks
ms.assetid: b226b376-9d8c-659a-9551-6341bbebed6f
ms.date: 06/08/2017
localization_priority: Normal
---


# SharedWorkspaceLinks object (Office)

A collection of the  **[SharedWorkspaceLink](Office.SharedWorkspaceLink.md)** objects in the current shared workspace.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Example

Use the  **[Links](Office.SharedWorkspace.Links.md)** property of the **[SharedWorkspace](Office.SharedWorkspace.md)** object to return a **SharedWorkspaceLinks** collection.


```vb
    Dim swsLinks As Office.SharedWorkspaceLinks 
    Set swsLinks = ActiveWorkbook.SharedWorkspace.Links 
    MsgBox "There are " &amp; swsLinks.Count &amp; _ 
        " link(s) in the current shared workspace.", _ 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsLinks = Nothing 

```


## Methods



|Name|
|:-----|
|[Add](Office.SharedWorkspaceLinks.Add.md)|

## Properties



|Name|
|:-----|
|[Application](Office.SharedWorkspaceLinks.Application.md)|
|[Count](Office.SharedWorkspaceLinks.Count.md)|
|[Creator](Office.SharedWorkspaceLinks.Creator.md)|
|[Item](Office.SharedWorkspaceLinks.Item.md)|
|[ItemCountExceeded](Office.SharedWorkspaceLinks.ItemCountExceeded.md)|
|[Parent](Office.SharedWorkspaceLinks.Parent.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)
