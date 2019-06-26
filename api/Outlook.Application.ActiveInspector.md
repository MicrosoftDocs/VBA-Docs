---
title: Application.ActiveInspector method (Outlook)
keywords: vbaol11.chm713
f1_keywords:
- vbaol11.chm713
ms.prod: outlook
api_name:
- Outlook.Application.ActiveInspector
ms.assetid: 3f2b6491-7b4b-8165-327e-b319711d5656
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ActiveInspector method (Outlook)

Returns the topmost  **[Inspector](Outlook.Inspector.md)** object on the desktop.


## Syntax

_expression_. `ActiveInspector`

_expression_ A variable that represents an **[Application](Outlook.Application.md)** object.


## Return value

An  **Inspector** that represents the topmost inspector on the desktop.


## Remarks

 Use this method to access the **Inspector** object that the user is most likely to be viewing.

If no inspector is active, returns  **Nothing**.


## Example

This Visual Basic for Applications (VBA) example uses the  **[ActiveInspector](Outlook.Application.ActiveInspector.md)** method to obtain the currently active **[Inspector](Outlook.Inspector.md)** object. The example saves and closes the item displayed in the active inspector without prompting the user. To run this example, you need to have an item displayed in an inspector window.


```vb
Sub CloseItem() 
 
 Dim myinspector As Outlook.Inspector 
 
 Dim myItem As Outlook.MailItem 
 
 
 
 Set myinspector = Application.ActiveInspector 
 
 Set myItem = myinspector.CurrentItem 
 
 myItem.Close olSave 
 
End Sub
```


## See also


[Application Object](Outlook.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
