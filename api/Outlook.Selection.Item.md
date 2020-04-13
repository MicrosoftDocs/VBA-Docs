---
title: Selection.Item method (Outlook)
keywords: vbaol11.chm86
f1_keywords:
- vbaol11.chm86
ms.prod: outlook
api_name:
- Outlook.Selection.Item
ms.assetid: 981b107a-14d7-2dd3-6449-2737b2801c3c
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Item method (Outlook)

Returns a Microsoft Outlook item or conversation header from the selection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a '[Selection](Outlook.Selection.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|Either the index number of the object, or a value used to match the default property of an object in the collection.|

## Return value

An **Object** that represents the specified item or conversation header.


## Remarks

Do not make any assumptions about the  **Item** method return type; your code should be able to handle multiple item types or a **[ConversationHeader](Outlook.ConversationHeader.md)** object. For example, the **Item** method can return an **[AppointmentItem](Outlook.AppointmentItem.md)**, **[MailItem](Outlook.MailItem.md)**, **[MeetingItem](Outlook.MeetingItem.md)**, or **[TaskItem](Outlook.TaskItem.md)** in the Inbox folder, depending on the value of the **[Selection.Location](Outlook.Selection.Location.md)** property.

The **[Selection](Outlook.Selection.md)** collection contains **ConversationHeader** objects only if you specify **olConversationHeaders** in the **[GetSelection](Outlook.Selection.GetSelection.md)** method of the **Selection** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example displays the sender of each selected item in the active explorer. It uses the  **[Count](Outlook.Selection.Count.md)** property and **[Item](Outlook.Selection.Item.md)** method of the **[Selection](Outlook.Selection.md)** object, returned by the **[Explorer.Selection](Outlook.Explorer.Selection.md)** property, to display the senders of all messages that are selected in the active explorer.


```vb
Sub GetSelectedItems() 
 Dim myOlExp As Outlook.Explorer 
 Dim myOlSel As Outlook.Selection 
 Dim mySender As Outlook.AddressEntry 
 Dim oMail As Outlook.MailItem 
 Dim oAppt As Outlook.AppointmentItem 
 Dim oPA As Outlook.PropertyAccessor 
 Dim strSenderID As String 
 Const PR_SENT_REPRESENTING_ENTRYID As String = _ 
 "http://schemas.microsoft.com/mapi/proptag/0x00410102" 
 Dim MsgTxt As String 
 Dim x As Long 
 
 MsgTxt = "Senders of selected items:" 
 Set myOlExp = Application.ActiveExplorer 
 Set myOlSel = myOlExp.Selection 
 For x = 1 To myOlSel.Count 
 If myOlSel.Item(x).Class = OlObjectClass.olMail Then 
 ' For mail item, use the SenderName property. 
 Set oMail = myOlSel.Item(x) 
 MsgTxt = MsgTxt & oMail.SenderName & ";" 
 ElseIf myOlSel.Item(x).Class = OlObjectClass.olAppointment Then 
 ' For appointment item, use the Organizer property. 
 Set oAppt = myOlSel.Item(x) 
 MsgTxt = MsgTxt & oAppt.Organizer & ";" 
 Else 
 ' For other items, use the property accessor to get sender ID, 
 ' then get the address entry to display the sender name. 
 Set oPA = myOlSel.Item(x).PropertyAccessor 
 strSenderID = oPA.GetProperty(PR_SENT_REPRESENTING_ENTRYID) 
 Set mySender = Application.Session.GetAddressEntryFromID(strSenderID) 
 MsgTxt = MsgTxt & mySender.Name & ";" 
 End If 
 Next x 
 Debug.Print MsgTxt 
End Sub
```


## See also


[Selection Object](Outlook.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
