---
title: NameSpace.CreateContactCard method (Outlook)
keywords: vbaol11.chm3536
f1_keywords:
- vbaol11.chm3536
ms.prod: outlook
api_name:
- Outlook.NameSpace.CreateContactCard
ms.assetid: d050e0e3-3c0d-bd01-f008-2628056625d1
ms.date: 06/08/2017
localization_priority: Normal
---


# NameSpace.CreateContactCard method (Outlook)

Creates an instance of a  **[ContactCard](Office.ContactCard.md)** object for the contact that is specified by the _AddressEntry_ parameter.


## Syntax

_expression_. `CreateContactCard`( `_Address_` )

_expression_ A variable that represents a '[NameSpace](Outlook.NameSpace.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _AddressEntry_|Required| **AddressEntry**|The  **[AddressEntry](Outlook.AddressEntry.md)** object that represents the user for whom the contact card is to be created.|

## Return value

The  **Office.ContactCard** object that is created for the specified user.


## Remarks

 The **ContactCard** object is available in the type library of Microsoft Office. Before calling **CreateContactCard** to create a contact card in Microsoft Outlook, Outlook must be logged into an Outlook session.

The  _AddressEntry_ parameter is an **AddressEntry** object that represents one of the following **AddressEntry** types defined in the **[OlAddressEntryUserType](Outlook.OlAddressEntryUserType.md)** enumeration:


- olExchangeDistributionListAddressEntry
    
- olExchangeRemoteUserAddressEntry
    
- olExchangeUserAddressEntry
    
- olOutlookContactAddressEntry
    
- olSmtpAddressEntry
    


Outlook raises the E_INVALIDARG error when you pass any of the following  **OlAddressEntryUserType** values as an argument to the **CreateContactCard** method:


- olExchangeAgentAddressEntry
    
- olExchangeOrganizationAddressEntry
    
- olExchangePublicFolderAddressEntry
    
- olLdapAddressEntry
    
- olOtherAddressEntry
    
- olOutlookDistributionListAddressEntry
    



## Example

 The following code sample in Microsoft Visual Basic for Applications (VBA) displays a Contact Card for the current user defined by the **[CurrentUser](Outlook.NameSpace.CurrentUser.md)** property of the **[NameSpace](Outlook.NameSpace.md)** object.

You cannot run this code directly from the VBA window. To run the code, click the  **Developer** tab, click the **Macros** menu, and then select **Project1.DisplayContactCardForCurrentUser**. For more information about the  **Developer** tab, see [Run in Developer Mode in Outlook](../outlook/How-to/Using-Visual-Basic-to-Customize-Outlook-Forms/run-in-developer-mode-in-outlook.md).




```vb
Sub DisplayContactCardForCurrentUser() 
 
 Dim oCC As Office.ContactCard 
 
 Dim oAddrEntry As Outlook.AddressEntry 
 
 Set oAddrEntry = Application.session.CurrentUser.AddressEntry 
 
 Set oCC = Application.session.CreateContactCard(oAddrEntry) 
 
 oCC.Show msoContactCardFull, 100, 100, 100, 100, 100, True 
 
End Sub
```


## See also


[NameSpace Object](Outlook.NameSpace.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]