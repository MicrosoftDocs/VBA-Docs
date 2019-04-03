---
title: MailItem.ReplyAll event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MailItem.ReplyAll
ms.assetid: f303adaf-71a3-e855-403d-2a6a3c8f9ceb
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.ReplyAll event (Outlook)

Occurs when the user selects the  **ReplyAll** action for an item, or when the **ReplyAll** method is called for the item, which is an instance of the parent object.


## Syntax

_expression_. `ReplyAll`( `_Response_` , `_Cancel_` )

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Response_|Required| **Object**|The new item being sent in response to the original message.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the reply all operation is not completed and the new item is not displayed.|

## Remarks

Returns the reply as a  **[MailItem](Outlook.MailItem.md)** object.


## Example

This Visual Basic for Applications (VBA) example uses the  **ReplyAll** event and reminds the user that proceeding will reply to all original recipients of an item and, depending on the user's response, either allows the action to continue or stops it. To use this example, open an existing mail item, run the `Initialize Handler()` procedure, then reply to the item.


```vb
Public WithEvents myItem As MailItem 
 
 
 
Sub Initialize_Handler() 
 
 Set myItem = Application.ActiveInspector.CurrentItem 
 
End Sub 
 
 
 
Private Sub myItem_ReplyAll(ByVal Response As Object, Cancel As Boolean) 
 
 Dim mymsg As String 
 
 Dim myResult As Integer 
 
 mymsg = "Do you really want to reply to all original recipients?" 
 
 myResult = MsgBox(mymsg, vbYesNo, "Flame Protector") 
 
 If myResult = vbNo Then 
 
 Cancel = True 
 
 End If 
 
End Sub
```


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]