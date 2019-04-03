---
title: SelectNamesDialog object (Outlook)
keywords: vbaol11.chm3156
f1_keywords:
- vbaol11.chm3156
ms.prod: outlook
api_name:
- Outlook.SelectNamesDialog
ms.assetid: 1522736a-3cad-9f1c-4da9-b52a3a01731c
ms.date: 06/08/2017
localization_priority: Normal
---


# SelectNamesDialog object (Outlook)

Displays the  **Select Names** dialog box for the user to select entries from one or more address lists, and returns the selected entries in the collection object specified by the property **[SelectNamesDialog.Recipients](Outlook.SelectNamesDialog.Recipients.md)**.


## Remarks

You can instantiate an instance of the  **SelectNamesDialog** object by calling **[NameSpace.GetSelectNamesDialog](Outlook.NameSpace.GetSelectNamesDialog.md)**.

The dialog box displayed by  **[SelectNamesDialog.Display](Outlook.SelectNamesDialog.Display.md)** is similar to the **Select Names** dialog box in the Outlook user interface. It observes the size and position settings of the built-in **Select Names** dialog box. However, its default state does not show **Message Recipients** above the **To**,  **Cc**, and  **Bcc** edit boxes. For more information on using the **SelectNamesDialog** object to display the **Select Names** dialog box, see [Display Names from the Address Book](../outlook/Concepts/Address-Book/display-names-from-the-address-book.md).


## Example

The following code sample shows how to use the  **SelectNamesDialog** object to display entries from the Contacts folder in a dialog box that resembles the **Select Names** dialog box in the Outlook user interface.


```vb
Sub ShowContactsInDialog() 
 
 Dim oDialog As SelectNamesDialog 
 
 Dim oAL As AddressList 
 
 Dim oContacts As Folder 
 
 
 
 Set oDialog = Application.Session.GetSelectNamesDialog 
 
 Set oContacts = _ 
 
 Application.Session.GetDefaultFolder(olFolderContacts) 
 
 
 
 'Look for the address list that corresponds with the Contacts folder 
 
 For Each oAL In Application.Session.AddressLists 
 
 If oAL.GetContactsFolder = oContacts Then 
 
 Exit For 
 
 End If 
 
 Next 
 
 With oDialog 
 
 'Initialize the dialog box with the address list representing the Contacts folder 
 
 .InitialAddressList = oAL 
 
 .ShowOnlyInitialAddressList = True 
 
 If .Display Then 
 
 'Recipients Resolved 
 
 'Access Recipients using oDialog.Recipients 
 
 End If 
 
 End With 
 
End Sub
```


## Methods



|Name|
|:-----|
|[Display](Outlook.SelectNamesDialog.Display.md)|
|[SetDefaultDisplayMode](Outlook.SelectNamesDialog.SetDefaultDisplayMode.md)|

## Properties



|Name|
|:-----|
|[AllowMultipleSelection](Outlook.SelectNamesDialog.AllowMultipleSelection.md)|
|[Application](Outlook.SelectNamesDialog.Application.md)|
|[BccLabel](Outlook.SelectNamesDialog.BccLabel.md)|
|[Caption](Outlook.SelectNamesDialog.Caption.md)|
|[CcLabel](Outlook.SelectNamesDialog.CcLabel.md)|
|[Class](Outlook.SelectNamesDialog.Class.md)|
|[ForceResolution](Outlook.SelectNamesDialog.ForceResolution.md)|
|[InitialAddressList](Outlook.SelectNamesDialog.InitialAddressList.md)|
|[NumberOfRecipientSelectors](Outlook.SelectNamesDialog.NumberOfRecipientSelectors.md)|
|[Parent](Outlook.SelectNamesDialog.Parent.md)|
|[Recipients](Outlook.SelectNamesDialog.Recipients.md)|
|[Session](Outlook.SelectNamesDialog.Session.md)|
|[ShowOnlyInitialAddressList](Outlook.SelectNamesDialog.ShowOnlyInitialAddressList.md)|
|[ToLabel](Outlook.SelectNamesDialog.ToLabel.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]