---
title: Explorer.ShowPane method (Outlook)
keywords: vbaol11.chm2776
f1_keywords:
- vbaol11.chm2776
ms.prod: outlook
api_name:
- Outlook.Explorer.ShowPane
ms.assetid: 3d2c9dd5-b660-e160-36db-73c23f95a7a2
ms.date: 06/08/2017
localization_priority: Normal
---


# Explorer.ShowPane method (Outlook)

Displays or hides a specific pane in the explorer.


## Syntax

_expression_. `ShowPane`( `_Pane_` , `_Visible_` )

_expression_ A variable that represents an **[Explorer](Outlook.Explorer.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Pane_|Required| **[OlPane](Outlook.OlPane.md)**|The pane to display.|
| _Visible_|Required| **Boolean**| **True** to make the pane visible, **False** to hide the pane.|

## Remarks




> [!NOTE] 
> You can also use the  **[Visible](Outlook.OutlookBarPane.Visible.md)** property of the **[OutlookBarPane](Outlook.OutlookBarPane.md)** object to display or hide the Outlook Bar.


## Example

This Microsoft Visual Basic for Applications (VBA) example uses the  **ShowPane** and **[IsPaneVisible](Outlook.Explorer.IsPaneVisible.md)** methods to hide the preview pane if it is visible or to display it if it is hidden.


```vb
Sub ShowHidePreviewPane() 
 
 Dim myOlExp As Outlook.Explorer 
 
 
 
 Set myOlExp = Application.ActiveExplorer 
 
 myOlExp.ShowPane olPreview, _ 
 
 Not myOlExp.IsPaneVisible(olPreview) 
 
End Sub
```


## See also


[Explorer Object](Outlook.Explorer.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]