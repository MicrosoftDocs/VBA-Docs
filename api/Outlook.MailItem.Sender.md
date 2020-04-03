---
title: MailItem.Sender property (Outlook)
keywords: vbaol11.chm3488
f1_keywords:
- vbaol11.chm3488
ms.prod: outlook
api_name:
- Outlook.MailItem.Sender
ms.assetid: c8afc3f8-fbf5-73b4-43f3-800e18aabb93
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.Sender property (Outlook)

Returns or sets an [AddressEntry](Outlook.AddressEntry.md) object that corresponds to the user of the account from which the [MailItem](Outlook.MailItem.md) is sent. Read/write.


## Syntax

_expression_. `Sender`

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Remarks

In a session where multiple accounts are defined in the profile, you can set this property to specify the account from which to send a mail item. Set this property to the  **AddressEntry** object of the user that is represented by the [CurrentUser](Outlook.Account.CurrentUser.md) property of a specific account.

If you set the  **Sender** property to an **AddressEntry** that does not have permissions to send messages on that account, Microsoft Outlook will raise an error.


## Example

Michael Bauer provided the following code example. Michael is a [Microsoft Most Valuable Professional](https://mvp.microsoft.com/) with expertise in developing Outlook solutions in Visual Basic and Visual Basic for Applications (VBA). Michael maintains a professional site at [VBOffice.net](https://www.vboffice.net/index.html?lang=en).

The following VBA code example shows how to display the details of the sender of an email. If the sender corresponds to a contact in the user's Outlook Contacts Address Book (CAB), the code example displays information about that contact in an inspector. If the sender is not a contact in the user's CAB, the code example displays details from the user's address entry (taken from the transport provider's address book container) in a dialog box. 

To display information about a sender, the user should have selected a  **MailItem** in the explorer. The code example also checks whether the selected **MailItem** has been sent, because the **Sender** property is defined only if the **Mailtem** has been sent. The example then accesses the **Sender** property to obtain the **AddressEntry** object that corresponds to the sender of that mail item, and displays the contact information, if it exists; otherwise, the example displays the address entry details.




```vb
 
Public Sub DisplaySenderDetails() 
 Dim Explorer As Outlook.Explorer 
 Dim CurrentItem As Object 
 Dim Sender As Outlook.AddressEntry 
 Dim Contact As Outlook.ContactItem 
 
 Set Explorer = Application.ActiveExplorer 
 
 ' Check whether any item is selected in the current folder. 
 If Explorer.Selection.Count Then 
 
 ' Get the first selected item. 
 Set CurrentItem = Explorer.Selection(1) 
 
 ' Check for the type of the selected item as only the 
 ' MailItem object has the Sender property. 
 If CurrentItem.Class = olMail Then 
 Set Sender = CurrentItem.Sender 
 
 ' There is no sender if the item has not been sent yet. 
 If Sender Is Nothing Then 
 MsgBox "There's no sender for the current email", vbInformation 
 Exit Sub 
 End If 
 
 Set Contact = Sender.GetContact 
 
 If Not Contact Is Nothing Then 
 ' The sender is stored in the contacts folder, 
 ' so the contact item can be displayed. 
 Contact.Display 
 
 Else 
 ' If the contact cannot be found, display the 
 ' address entry in the properties dialog box. 
 Sender.Details 0 
 End If 
 End If 
 End If 
End Sub
```


## See also


[MailItem Object](Outlook.MailItem.md)




[How to: Create a Sendable Item for a Specific Account Based on the Current Folder](../outlook/Concepts/Accounts/create-a-sendable-item-for-a-specific-account-based-on-the-current-folder-outloo.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
