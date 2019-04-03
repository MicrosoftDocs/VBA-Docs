---
title: Views.ViewAdd event (Outlook)
keywords: vbaol11.chm551
f1_keywords:
- vbaol11.chm551
ms.prod: outlook
api_name:
- Outlook.Views.ViewAdd
ms.assetid: 926eb4eb-7585-5bb0-b214-6e116a01375e
ms.date: 06/08/2017
localization_priority: Normal
---


# Views.ViewAdd event (Outlook)

Occurs when a view is added to the collection. Microsoft Outlook creates the new view and passes it to this event.


## Syntax

_expression_. `ViewAdd`( `_View_` )

_expression_ A variable that represents a [Views](Outlook.Views.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _View_|Required| **[View](Outlook.View.md)**|The new view added to the collection prior to this event.|

## Example

The following Microsoft Visual Basic for Applications (VBA) example displays the view's name and saves it when the  **ViewAdd** event is fired. Use the **[Save](Outlook.View.Save.md)** method after the properties have been modified to save the changes to the view. The sample code must be placed in a class module such as `ThisOutlookSession`, and the  `AddView()` procedure should be called before the event procedure can be called by Outlook.


```vb
Public WithEvents objViews As Outlook.Views 
 
 
 
Sub AddView() 
 
 Dim objView As Outlook.View 
 
 Set objViews = Application.ActiveExplorer.CurrentFolder.Views 
 
 Set objView = objViews.Add("Latest View1", olTableView, olViewSaveOptionAllFoldersOfType) 
 
End Sub 
 
 
 
Sub objViews_ViewAdd(ByVal View As View) 
 
'Displays name of new view 
 
 With View 
 
 Msgbox .Name & " was created programmatically." 
 
 .Save 
 
 End With 
 
End Sub
```


## See also


[Views Object](Outlook.Views.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]