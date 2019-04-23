---
title: Access Exchange User or Distribution List Information from the Address Book
ms.prod: outlook
ms.assetid: 077a8666-09c5-e641-0b9b-7d83133d931f
ms.date: 06/08/2017
localization_priority: Normal
---


# Access Exchange User or Distribution List Information from the Address Book

This topic describes the objects that support accessing information about an Exchange user or distribution list from the Address Book. 

The Address Book contains address lists of users, distribution lists, and other types of address entries, as enumerated by  **[OlAddressEntryUserType](../../../api/Outlook.OlAddressEntryUserType.md)**. Specifically, the Exchange user address entry and the Exchange distribution list address entry have many of their properties exposed as explicit built-in properties in the Outlook object model through the  **[ExchangeUser](../../../api/Outlook.ExchangeUser.md)** and **[ExchangeDistributionList](../../../api/Outlook.ExchangeDistributionList.md)** objects. Both of these objects inherit from the **[AddressEntry](../../../api/Outlook.AddressEntry.md)** object. They also support specific methods that facilitate accessing information about these entry types.

## Exchange User

The  **ExchangeUser** object supports properties like **[OfficeLocation](../../../api/Outlook.ExchangeUser.OfficeLocation.md)**,  **[JobTitle](../../../api/Outlook.ExchangeUser.JobTitle.md)**,  **[FirstName](../../../api/Outlook.ExchangeUser.FirstName.md)**, and  **[LastName](../../../api/Outlook.ExchangeUser.LastName.md)** that the parent **AddressEntry** object does not support. You can access these properties directly through the **ExchangeUser** object. You can access other properties of the Exchange user that are not exposed in the object model using **[ExchangeUser.PropertyAccessor](../../../api/Outlook.ExchangeUser.PropertyAccessor.md)**.

The  **ExchangeUser** object also supports methods like **[GetDirectReports](../../../api/Outlook.ExchangeUser.GetDirectReports.md)**,  **[GetExchangeUserManager](../../../api/Outlook.ExchangeUser.GetExchangeUserManager.md)**, and  **[GetMemberOfList](../../../api/Outlook.ExchangeUser.GetMemberOfList.md)** to facilitate accessing information specific to this Exchange user, such as full **AddressEntry** information for the associated direct reports, manager, and distribution lists.


## Security

Certain properties like  **OfficeLocation** and **JobTitle** are read-write and can only be updated (using **[ExchangeUser.Update](../../../api/Outlook.ExchangeUser.Update.md)**) by code that is running under an appropriate Exchange administrator account.


## Exchange Distribution List

 The **ExchangeDistributionList** object supports properties like **Alias**,  **[Comments](../../../api/Outlook.ExchangeDistributionList.Comments.md)**, and  **[PrimarySmtpAddress](../../../api/Outlook.ExchangeDistributionList.PrimarySmtpAddress.md)** that the parent **AddressEntry** object does not support. Other properties of the Exchange distribution list that are not exposed in the object model are accessible through **[ExchangeDistributionList.PropertyAccessor](../../../api/Outlook.ExchangeDistributionList.PropertyAccessor.md)**.

The  **ExchangeDistributionList** object also supports methods like **[GetExchangeDistributionListMembers](../../../api/Outlook.ExchangeDistributionList.GetExchangeDistributionListMembers.md)**,  **[GetMemberOfList](../../../api/Outlook.ExchangeDistributionList.GetMemberOfList.md)** and **[GetOwners](../../../api/Outlook.ExchangeDistributionList.GetOwners.md)** to facilitate accessing information specific to a distribution list, such as full **AddressEntry** information for the associated members in this distribution list, other distribution lists that this list is a member of, and owners of this list.


## Security

Certain properties like  **Comments** are read-write and can only be updated (using **[ExchangeDistributionList.Update](../../../api/Outlook.ExchangeDistributionList.Update.md)**) by code that is running under an appropriate Exchange administrator account.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]