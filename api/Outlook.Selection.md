---
title: Selection object (Outlook)
keywords: vbaol11.chm80
f1_keywords:
- vbaol11.chm80
ms.prod: outlook
api_name:
- Outlook.Selection
ms.assetid: 0b06a3ce-0445-db8f-e6e8-bb7bd469c50f
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection object (Outlook)

Contains the set of Outlook items currently selected in an explorer.


## Remarks

Use the  **[Selection](Outlook.Explorer.Selection.md)** property to return the **Selection** collection from the **[Explorer](Outlook.Explorer.md)** object.


## Example

The following example returns a **Selection** object from an **Explorer** object.


```vb
Set mySelectedItems = myExplorer.Selection
```


## Methods



|Name|
|:-----|
|[GetSelection](Outlook.Selection.GetSelection.md)|
|[Item](Outlook.Selection.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.Selection.Application.md)|
|[Class](Outlook.Selection.Class.md)|
|[Count](Outlook.Selection.Count.md)|
|[Location](Outlook.Selection.Location.md)|
|[Parent](Outlook.Selection.Parent.md)|
|[Session](Outlook.Selection.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)
[Selection Object Members](overview/Outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
