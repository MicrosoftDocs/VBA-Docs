---
title: Explorer.WindowState Property (Outlook)
keywords: vbaol11.chm2773
f1_keywords:
- vbaol11.chm2773
ms.prod: outlook
api_name:
- Outlook.Explorer.WindowState
ms.assetid: 787b6339-eb92-3ab6-df9f-82f6122facc5
ms.date: 06/08/2017
localization_priority: Normal
---


# Explorer.WindowState Property (Outlook)

Returns or sets the property with a constant in the  **[OlWindowState](Outlook.OlWindowState.md)** enumeration specifying the window state of an explorer or inspector window. Read/write.


## Syntax

_expression_. `WindowState`

_expression_ A variable that represents an [Explorer](./Outlook.Explorer.md) object.


## Example

This Microsoft Visual Basic for Applications example minimizes all open explorer windows. It uses the  **[Count](Outlook.Explorers.Count.md)** property and **[Item](Outlook.Explorers.Item.md)** method of the **[Explorers](Outlook.Explorers.md)** collection to enumerate the open explorer windows.


```vb
Sub MinimizeWindows() 
 
 Dim myOlExp As Outlook.Explorer 
 
 Dim myOlExps As Outlook.Explorers 
 
 
 
 Set myOlExps = Application.Explorers 
 
 For x = 1 To myOlExps.Count 
 
 myOlExps.Item(x).WindowState = olMinimized 
 
 Next x 
 
End Sub
```


## See also


[Explorer Object](Outlook.Explorer.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]