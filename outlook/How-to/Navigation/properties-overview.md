---
title: Properties Overview
ms.prod: outlook
ms.assetid: 242c9e89-a0c5-ff89-0d2a-410bd42a3461
ms.date: 10/10/2018
localization_priority: Normal
---


# Properties Overview

## Outlook properties

A property is an attribute of an Outlook object. Properties describe something about the object, such as the sender of a message item, or the number of items in a folder. Outlook defines many properties; these are referred to as built-in properties in this documentation. The Outlook object model exposes many built-in properties with string names, such as the **Subject** property of a mail item. 

These properties are further qualified as explicit built-in properties. Customers and service providers can extend the predefined properties of Outlook by creating new, custom properties. For example, through custom forms, customers can define properties to extend the functionality for a specific message class, and service providers can define properties to expose the unique features of their messaging system.

## Object model entry points

The Outlook object model provides several approaches to access Outlook properties, such as:

- Referencing a property directly from the parent object to access explicit built-in properties of item objects (for example, the **[MailItem.SenderEmailAddress](../../../api/Outlook.MailItem.SenderEmailAddress.md)** property)
    
- Using **[ItemProperties](../../../api/Outlook.ItemProperties.md)** and **[ItemProperty](../../../api/Outlook.ItemProperty.md)** to enumerate explicit built-in properties and custom properties and access custom properties of item objects
    
- Using **[UserProperties](../../../api/Outlook.UserProperties.md)** and **[UserProperty](../../../api/Outlook.UserProperty.md)** to enumerate and access custom properties of item objects
    
- Using the **[PropertyAccessor](../../../api/Outlook.PropertyAccessor.md)** object to access built-in and custom properties of the following objects:
    
  - **[AddressEntry](../../../api/Outlook.AddressEntry.md)**
    
  - **[AddressList](../../../api/Outlook.AddressList.md)**
    
  - **[Attachment](../../../api/Outlook.Attachment.md)**
    
  - **[ExchangeDistributionList](../../../api/Outlook.ExchangeDistributionList.md)**
    
  - **[ExchangeUser](../../../api/Outlook.ExchangeUser.md)**
    
  - **[Folder](../../../api/Outlook.Folder.md)**
    
  - **[Outlook item objects](../Items-Folders-and-Stores/outlook-item-objects.md)**
    
  - **[Recipient](../../../api/Outlook.Recipient.md)**
    
  - **[Store](../../../api/Outlook.Store.md)**
    

> [!NOTE]
> Although **ItemProperties** and **UserProperties** support enumerating explicit built-in properties, and **[UserProperties.Find](../../../api/Outlook.UserProperties.Find.md)** supports searching for explicit built-in properties, use these objects primarily for custom properties of item objects. Use the **PropertyAccessor** object to access properties of non-item objects, or item-level properties that are not explicitly exposed in the Outlook object model.

The following table shows when to use which entry points.

||Object.Property|UserProperty, UserProperties|ItemProperty, ItemProperties|PropertyAccessor|
|:-----|:-----|:-----|:-----|:-----|
|**Action on properties**|Get and set explicit built-in properties of item objects.|Enumerate, create, get, set, and remove custom properties of item objects.|Enumerate explicit built-in properties and custom properties of item objects; create, get, set, and remove custom properties of item objects.|Get and set built-in properties, and create, get, set, and remove custom properties. Objects include item objects and the following: **AddressEntry**, **AddressList**, **Attachment**, **ExchangeUser**, **ExchangeDistributionList**, **Folder**, **Recipient**, and **Store**. Access properties by the appropriate namespaces. For more information, see [Referencing Properties by Namespace](referencing-properties-by-namespace.md).|
|**Performance**|No performance overhead.|Enumerating and accessing properties using **UserProperties** can incur performance overhead.|Enumerating and accessing properties using **ItemProperties** can incur performance overhead.|Using the **PropertyAccessor** to access properties incurs performance overhead. For getting or setting multiple properties, use **GetProperties** and **SetProperties** as opposed to repeated calls to **GetProperty** and **SetProperty**.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]