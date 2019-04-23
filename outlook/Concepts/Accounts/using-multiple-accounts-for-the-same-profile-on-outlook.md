---
title: Using Multiple Accounts for the Same Profile on Outlook
ms.prod: outlook
ms.assetid: 9e06e076-d62a-37c8-4502-709da5a0b104
ms.date: 06/08/2017
localization_priority: Normal
---


# Using Multiple Accounts for the Same Profile on Outlook

You can sign in to Outlook by using a profile that specifies one or more accounts associated with different delivery stores. For a given session, the **NameSpace** object has members that maintain and access information for the primary Exchange account, and the [Accounts](../../../api/Outlook.NameSpace.Accounts.md) property of the [NameSpace](../../../api/Outlook.NameSpace.md) object holds information for all the accounts defined for the session's profile. 

The **NameSpace.Accounts** property returns an [Accounts](../../../api/Outlook.Accounts.md) collection for the current profile, tracking information for all accounts including Exchange, IMAP, and POP3 accounts, each of which can be associated with a different delivery store. To identify the primary Exchange account in the **Accounts** collection for a session, look for the account that has the [ExchangeStoreType](../../../api/Outlook.Store.ExchangeStoreType.md) property of the store (that is specified by **[Account.DeliveryStore](../../../api/Outlook.Account.DeliveryStore.md)**) equal to  **OlExchangeStoreType.olPrimaryExchangeMailbox**.

```vb
Account.DeliveryStore.ExchangeStoreType = OlExchangeStoreType.olPrimaryExchangeMailbox
```

The following table compares members of the **NameSpace** object and members of the [Account](../../../api/Outlook.Account.md), **Accounts**, or [Store](../../../api/Outlook.Store.md) object depending on whether the session's profile has one account or multiple accounts. If only the primary Exchange account is in the session's profile, use the following members of the NameSpace object. 

|**Description**|**Purpose**|
|:-----|:-----|
|Use the following members of the noted objects if there are multiple accounts in the session's profile.|**[AutoDiscoverConnectionMode](../../../api/Outlook.NameSpace.AutoDiscoverConnectionMode.md)** property, **[AutoDiscoverXml](../../../api/Outlook.NameSpace.AutoDiscoverXml.md)** property, **[AutoDiscoverComplete](../../../api/Outlook.NameSpace.AutoDiscoverComplete.md)** event|
|To use auto discovery of the Exchange server that hosts the primary Exchange account mailbox.|**[Account.AutoDiscoverConnectionMode](../../../api/Outlook.Account.AutoDiscoverConnectionMode.md)** property, **[Account.AutoDiscoverXml](../../../api/Outlook.Account.AutoDiscoverXml.md)** property, **[Accounts.AutoDiscoverComplete](../../../api/Outlook.Accounts.AutoDiscoverComplete.md)** event|
|To use auto discovery of the Exchange server that hosts the account mailbox.|**[ExchangeConnectionMode](../../../api/Outlook.NameSpace.ExchangeConnectionMode.md)** property, **[ExchangeMailboxServerName](../../../api/Outlook.NameSpace.ExchangeMailboxServerName.md)** property, **[ExchangeMailboxServerVersion](../../../api/Outlook.NameSpace.ExchangeMailboxServerVersion.md)** property|
|To obtain information for the Exchange server that hosts the primary Exchange account mailbox.|**[Account.ExchangeConnectionMode](../../../api/Outlook.Account.ExchangeConnectionMode.md)** property, **[Account.ExchangeMailboxServerName](../../../api/Outlook.Account.ExchangeMailboxServerName.md)** property, **[Account.ExchangeMailboxServerVersion](../../../api/Outlook.Account.ExchangeMailboxServerVersion.md)** property
|To obtain information for the Exchange server that hosts the account mailbox.|**[Categories](../../../api/Outlook.NameSpace.Categories.md)** property|
|To obtain a **[Categories](../../../api/Outlook.Categories.md)** collection that represents the Master Category List for the primary account of the session.|**[Store.Categories](../../../api/Outlook.Store.Categories.md)** property|
|To obtain a [Categories](../../../api/Outlook.Categories.md) collection that represents the categories defined for the store that is associated with an account in the session's profile.|**[CurrentUser](../../../api/Outlook.NameSpace.CurrentUser.md)** property|
|To obtain a **[Recipient](../../../api/Outlook.Recipient.md)** object that represents the user currently logged on for the session.|**[Account.CurrentUser](../../../api/Outlook.Account.CurrentUser.md)** property|
|To obtain a **Recipient** object that represents the user of the account that is defined in the session's profile. The account can be any account that Outlook supports including Exchange, IMAP, and POP3.|**[DefaultStore](../../../api/Outlook.NameSpace.DefaultStore.md)** property|
|To obtain the default store for the session's profile.| **[Account.DeliveryStore](../../../api/Outlook.Account.DeliveryStore.md)** property|
|To obtain the default delivery store for the account that is defined in the session's profile. The account can be any account that Outlook supports including Exchange, IMAP, and POP3.|**[GetAddressEntryFromID](../../../api/Outlook.NameSpace.GetAddressEntryFromID.md)** method|
|To obtain an **[AddressEntry](../../../api/Outlook.AddressEntry.md)** object that corresponds to the given entry ID.|**[Account.GetAddressEntryFromID](../../../api/Outlook.Account.GetAddressEntryFromID.md)** method|
|To obtain an **AddressEntry** object that corresponds to the account and given entry ID. The account can be any account that Outlook supports including Exchange, IMAP, and POP3.|**[GetRecipientFromID](../../../api/Outlook.NameSpace.GetRecipientFromID.md)** method|
|To obtain a **Recipient** object that corresponds to the given entry ID.|**[Account.GetRecipientFromID](../../../api/Outlook.Account.GetRecipientFromID.md)** method|
|To obtain a **Recipient** object that corresponds to the account and given entry ID. |The account can be any account that Outlook supports including Exchange, IMAP, and POP3.|

If you are operating with multiple accounts in the current profile, see the following tasks:

-  [How to: Obtain Information for Multiple Accounts](obtain-information-for-multiple-accounts.md)
    
-  [How to: Identify a Folder with an Account](identify-a-folder-with-an-account.md)
    
-  [How to: Create a Sendable Item for a Specific Account Based on the Current Folder](create-a-sendable-item-for-a-specific-account-based-on-the-current-folder-outloo.md)
    
-  [How to: Identify a Global Address List or a Set of Address Lists with a Store](../Address-Book/identify-the-global-address-list-or-a-set-of-address-lists-with-a-store.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]