---
title: Panes object (Outlook)
keywords: vbaol11.chm73
f1_keywords:
- vbaol11.chm73
ms.prod: outlook
api_name:
- Outlook.Panes
ms.assetid: 657d1adf-41e0-858f-c734-e435153ae9ad
ms.date: 06/08/2017
localization_priority: Normal
---


# Panes object (Outlook)

Contains the panes displayed by the specified  **[Explorer](Outlook.Explorer.md)**.


## Remarks

Use the  **[Panes](Outlook.Explorer.Panes.md)** property to return the **Panes** collection object from an **Explorer** object.

Use the  **[Item](Outlook.Panes.Item.md)** method to retrieve a specific pane.

For Microsoft Outlook 2000 and later, the  **Shortcuts** pane is the only pane that you can access through the **Panes** object.


## Example

The following Visual Basic for Applications (VBA) example retrieves the  **Panes** object from an **Explorer** object.


```vb
Set myPanes = myExplorer.Panes
```

The following example retrieves the  **[OutlookBarPane](Outlook.OutlookBarPane.md)** object representing the **Shortcuts** pane.




```vb
Set myOLBarPane = myExplorer.Panes.Item("OutlookBar") 

```


## Methods



|Name|
|:-----|
|[Item](Outlook.Panes.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.Panes.Application.md)|
|[Class](Outlook.Panes.Class.md)|
|[Count](Outlook.Panes.Count.md)|
|[Parent](Outlook.Panes.Parent.md)|
|[Session](Outlook.Panes.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]