---
title: OutlookBarPane.Visible property (Outlook)
keywords: vbaol11.chm366
f1_keywords:
- vbaol11.chm366
ms.prod: outlook
api_name:
- Outlook.OutlookBarPane.Visible
ms.assetid: d9d00e7a-52ef-b481-7a56-729e1ac04534
ms.date: 06/08/2017
localization_priority: Normal
---


# OutlookBarPane.Visible property (Outlook)

Returns or sets a **Boolean** indicating the visible state of the specified object. Read/write.


## Syntax

_expression_.**Visible**

_expression_ A variable that represents an [OutlookBarPane](Outlook.OutlookBarPane.md) object.


## Remarks

 **True** to display the object; **False** to hide the object.

You can also use the  **[ShowPane](Outlook.Explorer.ShowPane.md)** method or the **[IsPaneVisible](Outlook.Explorer.IsPaneVisible.md)** method of an **[Explorer](Outlook.Explorer.md)** object to set or retrieve this value.


## Example

This Microsoft Visual Basic for Applications (VBA) example toggles the visible state of the Shortcuts pane.


```vb
Sub ShowHideShortcutsBar() 
 
 Dim myOlBar As Outlook.OutlookBarPane 
 
 
 
 Set myOlBar = Application.ActiveExplorer.Panes.Item("OutlookBar") 
 
 myOlBar.Visible = Not myOlBar.Visible 
 
End Sub
```


## See also


[OutlookBarPane Object](Outlook.OutlookBarPane.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]