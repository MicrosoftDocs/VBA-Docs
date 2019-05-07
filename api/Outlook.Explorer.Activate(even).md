---
title: Explorer.Activate event (Outlook)
keywords: vbaol11.chm449
f1_keywords:
- vbaol11.chm449
ms.prod: outlook
api_name:
- Outlook.Explorer.Activate
ms.assetid: 8543d347-baf5-cdc9-2366-11c9917e035e
ms.date: 06/08/2017
localization_priority: Normal
---


# Explorer.Activate event (Outlook)

Occurs when an explorer becomes the active window, either as a result of user action or through program code.


## Syntax

_expression_.**Activate**

_expression_ A variable that represents an **[Explorer](Outlook.Explorer.md)** object.


## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

This code example uses the  **[WindowState](Outlook.Explorer.WindowState.md)** property to maximize the topmost explorer window when the **Activate** event occurs. The sample code must be placed in a class module, and the `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Public WithEvents myOlExp As Outlook.Explorer 
 
 
 
Public Sub Initialize_handler() 
 
 Set myOlExp = Application.ActiveExplorer 
 
End Sub 
 
 
 
Private Sub myOlExp_Activate() 
 
 If myOlExp.WindowState = olNormalWindow Then _ 
 
 myOlExp.WindowState = olMaximized 
 
End Sub
```


## See also


[Explorer Object](Outlook.Explorer.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]