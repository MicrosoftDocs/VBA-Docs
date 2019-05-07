---
title: Explorer.IsPaneVisible method (Outlook)
keywords: vbaol11.chm2775
f1_keywords:
- vbaol11.chm2775
ms.prod: outlook
api_name:
- Outlook.Explorer.IsPaneVisible
ms.assetid: d547978a-f6b4-06ea-2358-8b6a81230240
ms.date: 06/08/2017
localization_priority: Normal
---


# Explorer.IsPaneVisible method (Outlook)

Returns a  **Boolean** indicating whether a specific explorer pane is visible.


## Syntax

_expression_. `IsPaneVisible`( `_Pane_` )

_expression_ A variable that represents an **[Explorer](Outlook.Explorer.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Pane_|Required| **[OlPane](Outlook.OlPane.md)**|The pane to check.|

## Return value

 **True** if the specified pane is displayed in the explorer; otherwise, **False**.


## Remarks

You can also use the  **[Visible](Outlook.OutlookBarPane.Visible.md)** property of the **[OutlookBarPane](Outlook.OutlookBarPane.md)** object to determine whether the **Shortcuts** pane is visible.


## Example

This Microsoft Visual Basic for Applications (VBA) sample uses the  **IsPaneVisible** method to determine whether the preview pane is visible and uses the **[ShowPane](Outlook.Explorer.ShowPane.md)** method to display it if it is not visible. Use the **olNavigationPane** constant to hide or display the navigation pane.


```vb
Sub HidePreviewPane() 
 
 Dim myOlExp As Outlook.Explorer 
 
 Set myOlExp = Application.ActiveExplorer 
 
 If myOlExp.IsPaneVisible(olPreview) = False Then 
 
 myOlExp.ShowPane olPreview, True 
 
 End If 
 
 Set myOlExp = Nothing 
 
End Sub
```


## See also


[Explorer Object](Outlook.Explorer.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]