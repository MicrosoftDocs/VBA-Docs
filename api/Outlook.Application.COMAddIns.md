---
title: Application.COMAddIns property (Outlook)
keywords: vbaol11.chm719
f1_keywords:
- vbaol11.chm719
ms.prod: outlook
api_name:
- Outlook.Application.COMAddIns
ms.assetid: f911199d-dc2e-9b88-d807-a5737a39f29e
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.COMAddIns property (Outlook)

Returns a  **COMAddIns** collection that represents all the Component Object Model (COM) add-ins currently loaded in Microsoft Outlook.


## Syntax

_expression_.**COMAddIns**

_expression_ A variable that represents an **[Application](Outlook.Application.md)** object.


## Example

This Microsoft Visual Basic for Applications (VBA) example displays the number of COM add-ins currently loaded.


```vb
Private Sub CountCOMAddins() 
 
 MsgBox "There are " & _ 
 
 Application.COMAddIns.Count & " COM add-ins." 
 
End Sub
```


## See also


[Application Object](Outlook.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]