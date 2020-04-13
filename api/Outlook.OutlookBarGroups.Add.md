---
title: OutlookBarGroups.Add method (Outlook)
keywords: vbaol11.chm352
f1_keywords:
- vbaol11.chm352
ms.prod: outlook
api_name:
- Outlook.OutlookBarGroups.Add
ms.assetid: cf3e449f-82c2-463b-1b30-c7a0729d9208
ms.date: 06/08/2017
localization_priority: Normal
---


# OutlookBarGroups.Add method (Outlook)

Adds a new, empty group to the  **Shortcuts** pane.


## Syntax

_expression_.**Add** (_Name_, _Index_)

_expression_ A variable that represents an [OutlookBarGroups](Outlook.OutlookBarGroups.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the group being created.|
| _Index_|Optional| **Long**|The position at which the new group will be inserted in the  **Shortcuts** pane. Position one is at the top of the bar.|

## Return value

An **[OutlookBarGroup](Outlook.OutlookBarGroup.md)** object that represents the new group.


## Example

This Microsoft Visual Basic for Applications (VBA) example adds a group named Marketing as the last group in the  **Shortcuts** pane.


```vb
Sub AddGroup() 
 Dim myolBar As Outlook.OutlookBarPane 
 
 Set myolBar = Application.ActiveExplorer.Panes.Item("OutlookBar") 
 myolBar.Contents.Groups.Add "Marketing", myolBar.Contents.Groups.Count + 1 
End Sub
```


## See also


[OutlookBarGroups Object](Outlook.OutlookBarGroups.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]