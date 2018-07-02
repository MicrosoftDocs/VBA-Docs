---
title: SharedWorkspaceFiles Object (Office)
keywords: vbaof11.chm267000
f1_keywords:
- vbaof11.chm267000
ms.prod: office
api_name:
- Office.SharedWorkspaceFiles
ms.assetid: 5e2937f7-f794-dffb-a1ec-69ea9a9e3546
ms.date: 06/08/2017
---


# SharedWorkspaceFiles Object (Office)

A collection of the  **[SharedWorkspaceFile](Office.SharedWorkspaceFile.md)** objects in the current shared workspace.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Example

Use the  **[Files](Office.SharedWorkspace.Files.md)** property of the **[SharedWorkspace](Office.SharedWorkspace.md)** object to return a **SharedWorkspaceFiles** collection.


```vb
    Dim swsFiles As Office.SharedWorkspaceFiles 
    Set swsFiles = ActiveWorkbook.SharedWorkspace.Files 
    MsgBox "There are " &amp; swsFiles.Count &amp; _ 
        " file(s) 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsFiles = Nothing 

```


## Methods



|**Name**|
|:-----|
|[Add](Office.SharedWorkspaceFiles.Add.md)|

## Properties



|**Name**|
|:-----|
|[Application](Office.SharedWorkspaceFiles.Application.md)|
|[Count](Office.SharedWorkspaceFiles.Count.md)|
|[Creator](Office.SharedWorkspaceFiles.Creator.md)|
|[Item](Office.SharedWorkspaceFiles.Item.md)|
|[ItemCountExceeded](Office.SharedWorkspaceFiles.ItemCountExceeded.md)|
|[Parent](Office.SharedWorkspaceFiles.Parent.md)|

## See also





[Object Model Reference](./overview/reference-object-library-reference-for-office.md)
