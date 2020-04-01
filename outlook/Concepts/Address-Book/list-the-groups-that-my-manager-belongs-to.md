---
title: List the Groups that My Manager Belongs to
ms.prod: outlook
ms.assetid: 2f0ff92c-e026-4f62-c039-fbda9aaf1546
ms.date: 06/08/2017
localization_priority: Normal
---


# List the Groups that My Manager Belongs to

This topic describes how to obtain the names of the Exchange distribution lists that the manager of the current user belongs to. It uses the **[ExchangeUser](../../../api/Outlook.ExchangeUser.md)** object to obtain specific Exchange user information such as the user's Exchange account alias, details about the user's manager, and the distribution lists that the user's manager has joined:


1. Obtain the current user's ExchangeUser object. Use the **[GetExchangeUser](../../../api/Outlook.AddressEntry.GetExchangeUser.md)** method of the **[AddressEntry](../../../api/Outlook.AddressEntry.md)** object for the current user to get the **ExchangeUser** object that represents the current user.
    
2. Obtain the distribution lists that the user's manager has joined.Use the **ExchangeUser** methods **[GetExchangeUserManager](../../../api/Outlook.ExchangeUser.GetExchangeUserManager.md)** and **[GetMemberOfList](../../../api/Outlook.ExchangeUser.GetMemberOfList.md)** to find these distribtution lists. Use the **[ExchangeDistributionList](../../../api/Outlook.ExchangeDistributionList.md)** object to obtain further information about a distribution list, such as its display name.
    

```vb
Sub ShowManagerDistLists() 
 Dim oAE As Outlook.AddressEntry 
 Dim oExUser As Outlook.ExchangeUser 
 Dim oDistListEntries As Outlook.AddressEntries 
 
 'Obtain the AddressEntry for CurrentUser 
 Set oExUser = _ 
 Application.Session.CurrentUser.AddressEntry.GetExchangeUser 
 
 'Obtain distribution lists that the user's manager has joined 
 Set oDistListEntries = oExUser.GetExchangeUserManager.GetMemberOfList 
 For Each oAE In oDistListEntries 
 If oAE.AddressEntryUserType = _ 
 olExchangeDistributionListAddressEntry Then 
 Debug.Print (oAE.name) 
 End If 
 Next 
End Sub 
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]