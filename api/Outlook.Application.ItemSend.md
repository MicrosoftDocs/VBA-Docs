---
title: Application.ItemSend event (Outlook)
keywords: vbaol11.chm429
f1_keywords:
- vbaol11.chm429
ms.prod: outlook
api_name:
- Outlook.Application.ItemSend
ms.assetid: 54f506ea-87a2-29b9-2b33-67bc87167933
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ItemSend event (Outlook)

Occurs whenever an Microsoft Outlook item is sent, either by the user through an  **[Inspector](Outlook.Inspector.md)** (before the inspector is closed, but after the user clicks the **Send** button) or when the **[Send](Outlook.MailItem.Send(method).md)** method for an Outlook item, such as **[MailItem](Outlook.MailItem.md)**, is used in a program.


## Syntax

_expression_. `ItemSend`( `_Item_` , `_Cancel_` )

_expression_ A variable that represents an **[Application](Outlook.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The item being sent.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the send action is not completed and the inspector is left open.|

## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

The following Microsoft Visual Basic for Applications (VBA) example shows how to cancel the  **ItemSend** event in response to user input. The sample code must be placed in a class module, and the `Initialize_handler` routine must be called before the event procedure can be called by Outlook.


```vb
Public WithEvents myOlApp As Outlook.Application 
 
 
 
Public Sub Initialize_handler() 
 
 Set myOlApp = Outlook.Application 
 
End Sub 
 
 
 
Private Sub myOlApp_ItemSend(ByVal Item As Object, Cancel As Boolean) 
 
 Dim prompt As String 
 
 prompt = "Are you sure you want to send " & Item.Subject & "?" 
 
 If MsgBox(prompt, vbYesNo + vbQuestion, "Sample") = vbNo Then 
 
 Cancel = True 
 
 End If 
 
End Sub
```


## See also


[Application Object](Outlook.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
