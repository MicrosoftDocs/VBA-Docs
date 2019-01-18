---
title: Creating a Property
ms.prod: outlook
ms.assetid: 511754a6-e9b7-6ad6-7159-62105ec53a76
ms.date: 06/08/2017
localization_priority: Normal
---


# Creating a Property

Outlook provides several ways to add custom properties.

|ObjectProperty|[UserProperties.Add](../../../api/Outlook.UserProperties.Add.md)|[ItemProperties.Add](../../../api/Outlook.ItemProperties.Add.md)|[PropertyAccessor.SetProperty](../../../api/Outlook.PropertyAccessor.SetProperty.md)|[PropertyAccessor.SetProperties](../../../api/Outlook.PropertyAccessor.SetProperties.md)|
|:-----|:-----|:-----|:-----|:-----|
|**Action**|Adds a custom property specified by _Name_ and _Type_. If a property of the same name and type already exists, it will be overwritten by a new property. The default value for _AddToFolderFields_ allows adding the property to the item and as a view field to the folder.|Adds a custom property specified by _Name_ and _Type_ even if a property of the same name and type already exists. The default value for _AddToFolderFields_ allows adding the property to the item and as a view field to the folder.|Adds a custom property specified by _SchemaName_ if the provider and the parent object support property creation, the property does not already exist, and a valid schema name is specified for the property.|For each property in _SchemaNames_, **[PropertyAccessor.SetProperties](../../../api/Outlook.PropertyAccessor.SetProperties.md)** adds it as a custom property if the provider and the parent object support property creation, the property does not already exist, and a valid schema name is specified for the property.|
|**Applicable objects**|All [Outlook item objects](../Items-Folders-and-Stores/outlook-item-objects.md) except Office document items (**[DocumentItem](../../../api/Outlook.DocumentItem.md)** objects).|All Outlook item objects except Office document items (**DocumentItem** objects).|All Outlook item objects including **DocumentItem** objects.|All Outlook item objects including **DocumentItem** objects.|
| **Property initial value**| **Empty** in VBA; requires subsequent assignment.| **Empty** in VBA; requires subsequent assignment.|Specified by _Value_.|Specified by the value of the corresponding element in the _Values_ array.|
| **Property type**|Specified by _Type_.|Specified by _Type_.|If the property is specified by the MAPI proptag or id namespace, the property type is contained in the lowest 16 bits of the identifier; otherwise, the property type is determined by the type of  _Value_.|Type of each property is determined by the same principles as in the **SetProperty** column; where the property is not specified by any namespace involving its proptag, its property type is the type of the corresponding element in the _Values_ array.|
| **Upon property change**|The **CustomPropertyChange** event will fire on property change.|The **CustomPropertyChange** event will fire on property change.|An item-level property added this way does not become part of the item's **[UserProperties](../../../api/Outlook.UserProperties.md)** collection. It will not generate Outlook Object Model events when it is changed.|Same event considerations as in the  **SetProperty** column.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]