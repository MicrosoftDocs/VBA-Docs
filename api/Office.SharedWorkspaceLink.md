---
title: SharedWorkspaceLink object (Office)
keywords: vbaof11.chm270000
f1_keywords:
- vbaof11.chm270000
ms.prod: office
api_name:
- Office.SharedWorkspaceLink
ms.assetid: eb36dbed-fc41-08df-3cbc-affbaf5f9784
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspaceLink object (Office)

Represents a URL link saved in a shared document workspace site.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Remarks

Use the **SharedWorkspaceLink** object to manage links to additional documents and information of interest to the members who are collaborating on the documents in the shared workspace site.

Use the **Item** (_index_) property of the **SharedWorkspaceLinks** collection to return a specific **SharedWorkspaceLink** object.

Use the **Description** property to set the link description that appears on the **Links** tab of the **Shared Workspace** pane and on the workspace webpage. Use the **URL** property to set the destination address of the link. Use the **Notes** property to supply additional information about the link.

Use the **Save** method to upload changes to the server after you modify properties of the **SharedWorkspaceLink** object.

Use the **CreatedBy**, **CreatedDate**, **ModifiedBy**, and **ModifiedDate** properties to return information about the history of each link.


## Example

The following example modifies the first link in the shared workspace site to point to the Microsoft Developer Network home page, and then uploads the changes to the server.


```vb
    Dim swsLink As Office.SharedWorkspaceLink 
    Set swsLink = ActiveWorkbook.SharedWorkspace.Links(1) 
    With swsLink 
        .Description = "MSDN Home Page" 
        .URL = "https://msdn.microsoft.com/" 
        .Notes = "My favorite site for developers!" 
        .Save 
    End With 
    Set swsLink = Nothing 

```

## See also

- [SharedWorkspaceLink object members](overview/Library-Reference/sharedworkspacelink-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]