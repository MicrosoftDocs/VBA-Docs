---
title: Deleting a Property
ms.prod: outlook
ms.assetid: 69d97b27-f60e-6c7a-36c8-a10986101219
ms.date: 06/08/2017
localization_priority: Normal
---


# Deleting a Property

Outlook provides several ways to remove a custom property.

|_ObjectProperty_|[UserProperties.Remove](../../../api/Outlook.UserProperties.Remove.md)|[ItemProperties.Remove](../../../api/Outlook.ItemProperties.Remove.md)|[PropertyAccessor.DeleteProperty](../../../api/Outlook.PropertyAccessor.DeleteProperty.md)|[PropertyAccessor.DeleteProperties](../../../api/Outlook.PropertyAccessor.DeleteProperties.md)|
|:-----|:-----|:-----|:-----|:-----|
|**Action**|Removes a custom property specified by _Index_ in the **[UserProperties](../../../api/Outlook.UserProperties.md)** collection for the item. The **UserProperties** collection is one-based.|Removes a custom property specified by _Index_ in the **[ItemProperties](../../../api/Outlook.ItemProperties.md)** collection for the item. The **ItemProperties** collection is zero-based. You can only remove custom properties in the collection, and they are denoted by **[IsUserProperty](../../../api/Outlook.ItemProperty.IsUserProperty.md)**. You cannot remove explicit built-in properties.|Removes a custom property specified by _SchemaName_, provided that the property is not read-only and the caller has permission to delete the property (for example, the caller is the owner of the folder to which the property has been added). You cannot remove a built-in Outlook or MAPI property.|For each custom property in _SchemaNames_, removes it provided that the same conditions described in the **PropertyAccessor.DeleteProperty** column are true. Any error will be returned in the corresponding element in the resultant error array.|
|**Applicable objects**|All [Outlook item objects](../Items-Folders-and-Stores/outlook-item-objects.md) except Office document items (**[DocumentItem](../../../api/Outlook.DocumentItem.md)** objects).|All Outlook item objects except Office document items (**DocumentItem** objects).|All Outlook item objects excluding the **DocumentItem** object, and any of the following objects: **[AddressEntry](../../../api/Outlook.AddressEntry.md)**, **[AddressList](../../../api/Outlook.AddressList.md)**, **[Attachment](../../../api/Outlook.Attachment.md)**, **[ExchangeDistributionList](../../../api/Outlook.ExchangeDistributionList.md)**, **[ExchangeUser](../../../api/Outlook.ExchangeUser.md)**, **[Folder](../../../api/Outlook.Folder.md)**, **[Recipient](../../../api/Outlook.Recipient.md)**, and **[Store](../../../api/Outlook.Store.md)** objects.|Same objects as listed in the **DeleteProperty** column.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]