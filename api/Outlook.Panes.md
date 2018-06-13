---
title: Panes Object (Outlook)
keywords: vbaol11.chm73
f1_keywords:
- vbaol11.chm73
ms.prod: outlook
api_name:
- Outlook.Panes
ms.assetid: 657d1adf-41e0-858f-c734-e435153ae9ad
ms.date: 06/08/2017
---


# Panes Object (Outlook)

Contains the panes displayed by the specified  **[Explorer](Outlook.Explorer.md)**.


## Remarks

Use the  **[Panes](Outlook.Explorer.Panes.md)** property to return the **Panes** collection object from an **Explorer** object.

Use the  **[Item](Outlook.Panes.Item.md)** method to retrieve a specific pane.

For Microsoft Outlook 2000 and later, the  **Shortcuts** pane is the only pane that you can access through the **Panes** object.


## Example

The following Visual Basic for Applications (VBA) example retrieves the  **Panes** object from an **Explorer** object.


```
Set myPanes = myExplorer.Panes
```

The following example retrieves the  **[OutlookBarPane](Outlook.OutlookBarPane.md)** object representing the **Shortcuts** pane.




```
Set myOLBarPane = myExplorer.Panes.Item("OutlookBar") 

```


## Methods



|**Name**|
|:-----|
|[Item](Outlook.Panes.Item.md)|

## Properties



|**Name**|
|:-----|
|[Application](Outlook.Panes.Application.md)|
|[Class](Outlook.Panes.Class.md)|
|[Count](Outlook.Panes.Count.md)|
|[Parent](Outlook.Panes.Parent.md)|
|[Session](panes-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
