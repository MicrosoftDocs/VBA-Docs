---
title: Application.Inspectors property (Outlook)
keywords: vbaol11.chm721
f1_keywords:
- vbaol11.chm721
ms.prod: outlook
api_name:
- Outlook.Application.Inspectors
ms.assetid: c2dde847-d033-90e3-30d2-62ff375d6843
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Inspectors property (Outlook)

Returns an  **[Inspectors](Outlook.Inspectors.md)** collection object that contains the **[Inspector](Outlook.Inspector.md)** objects representing all open inspectors. Read-only.


## Syntax

_expression_. `Inspectors`

_expression_ A variable that represents an **[Application](Outlook.Application.md)** object.


## Example

This Microsoft Visual Basic example uses the  **[Inspectors](Outlook.Application.Inspectors.md)** property and the **[Count](Outlook.Inspectors.Count.md)** property and **[Item](Outlook.Inspectors.Item.md)** method of the **[Inspectors](Outlook.Inspectors.md)** object to display the captions of all inspector windows.


```vb
Private Sub CommandButton1_Click() 
 
 Dim myInspectors As Outlook.Inspectors 
 
 Dim x as Integer 
 
 Dim iCount As Integer 
 
 
 
 Set myInspectors = Application.Inspectors 
 
 iCount = Application.Inspectors.Count 
 
 If iCount > 0 Then 
 
 For x = 1 To iCount 
 
 MsgBox myInspectors.Item(x).Caption 
 
 Next x 
 
 Else 
 
 MsgBox "No inspector windows are open." 
 
 End If 
 
End Sub
```


## See also


[Application Object](Outlook.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]