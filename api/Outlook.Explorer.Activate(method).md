---
title: Explorer.Activate method (Outlook)
keywords: vbaol11.chm2774
f1_keywords:
- vbaol11.chm2774
ms.prod: outlook
api_name:
- Outlook.Explorer.Activate
ms.assetid: 53f33d64-7a33-6772-4abc-fe328d3abb57
ms.date: 06/08/2017
localization_priority: Normal
---


# Explorer.Activate method (Outlook)

Activates an explorer window by bringing it to the foreground and setting keyboard focus.


## Syntax

_expression_.**Activate**

_expression_ A variable that represents an **[Explorer](Outlook.Explorer.md)** object.


## Example

This Microsoft Visual Basic for Applications example responds to the  **[NewMail](Outlook.Application.NewMail.md)** event by activating the explorer window. The sample code must be placed in a class module, and the `Initialize_handlers` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Public WithEvents myOlExp As Outlook.Explorer 
 
Public Sub Initialize_handlers() 
 Set myOlExp = Application.ActiveExplorer 
End Sub 
 
Private Sub NewMail() 
 myOlExp.Activate 
End Sub
```


## See also


[Explorer Object](Outlook.Explorer.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]