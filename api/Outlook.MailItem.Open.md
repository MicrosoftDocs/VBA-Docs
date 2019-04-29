---
title: MailItem.Open event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MailItem.Open
ms.assetid: 656c16f7-d561-a8f7-e859-9ac24f357769
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.Open event (Outlook)

Occurs when an instance of the parent object is being opened in an  **[Inspector](Outlook.Inspector.md)**.


## Syntax

_expression_.**Open** (_Cancel_)

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the open operation is not completed and the inspector is not displayed.|

## Remarks

When this event occurs, the  **Inspector** object is initialized but not yet displayed. The **Open** event differs from the **[Read](Outlook.AppointmentItem.Read.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an inspector.

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the open operation is not completed and the inspector is not displayed.


## Example

This Visual Basic for Applications (VBA) example uses the  **Open** event to display the "All Fields" page every time the item is opened.


```vb
Public WithEvents myItem As Outlook.MailItem 
 
 
 
Sub Initialize_handler() 
 
 Set myItem = Application.Session.GetDefaultFolder(olFolderInbox).Items(1) 
 
 myItem.Display 
 
End Sub 
 
 
 
Private Sub myItem_Open(Cancel As Boolean) 
 
 myItem.GetInspector.SetCurrentFormPage "All Fields" 
 
End Sub
```

This Visual Basic for Applications example uses the  **[Unread](Outlook.MailItem.UnRead.md)** property to detect whether the item has been previously read. If it has, then it asks if the user wants to open it. If the user answers No, the return value is set to **False** to prevent the item from opening.




```vb
Public WithEvents myItem As Outlook.MailItem 
 
 
 
Sub Initialize_handler() 
 
 Set myItem = Application.Session.GetDefaultFolder(olFolderInbox).Items(1) 
 
 myItem.Display 
 
End Sub 
 
 
 
Private Sub myItem_Open(Cancel As Boolean) 
 
 Dim mymsg As String 
 
 If myItem.UnRead = False Then 
 
 mymsg = "You have already read this message. Do you want to open this message again?" 
 
 If MsgBox(mymsg, 4) = 6 Then 
 
 Cancel = False 
 
 Else 
 
 Cancel = True 
 
 End If 
 
 End If 
 
End Sub
```


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
