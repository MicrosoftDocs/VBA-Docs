---
title: UserProperties object (Outlook)
keywords: vbaol11.chm202
f1_keywords:
- vbaol11.chm202
ms.prod: outlook
api_name:
- Outlook.UserProperties
ms.assetid: 20b49c86-d74f-9bda-382c-559af278c148
ms.date: 06/08/2017
localization_priority: Normal
---


# UserProperties object (Outlook)

Contains  **[UserProperty](Outlook.UserProperty.md)** objects that represent the custom properties of an Outlook item.


## Remarks

Use the  **UserProperties** property to return the **UserProperties** object for an Outlook item. This applies to all Outlook items except for the **[NoteItem](Outlook.NoteItem.md)**.

Use the  **[Add](Outlook.UserProperties.Add.md)** method to create a new **UserProperty** for an item and add it to the **UserProperties** object. The **Add** method allows you to specify a name and type for the new property. When you create a new property, it can also be added as a custom field to the folder that contains the item (using the same name as the property) by setting the _AddToFolderFields_ parameter to **True** when calling the **Add** method. That field can then be used as a column in folder views.

Use  **UserProperties** (_index_), where _index_ is a name or one-based index number, to return a single **[UserProperty](Outlook.UserProperty.md)** object.

You can use the  **[UserDefinedProperties](Outlook.Folder.UserDefinedProperties.md)** property of the **[Folder](Outlook.Folder.md)** object to retrieve and examine the definitions of custom item-level properties that a folder can display in a view.

To get or set multiple custom properties, use the  **[PropertyAccessor](Outlook.PropertyAccessor.md)** object instead of the **UserProperties** object for better performance.


## Example

The following example adds a custom text property named MyPropName to myItem.


```vb
Set myProp = myItem.UserProperties.Add("MyPropName", olText)
```


## Methods



|Name|
|:-----|
|[Add](Outlook.UserProperties.Add.md)|
|[Find](Outlook.UserProperties.Find.md)|
|[Item](Outlook.UserProperties.Item.md)|
|[Remove](Outlook.UserProperties.Remove.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.UserProperties.Application.md)|
|[Class](Outlook.UserProperties.Class.md)|
|[Count](Outlook.UserProperties.Count.md)|
|[Parent](Outlook.UserProperties.Parent.md)|
|[Session](Outlook.UserProperties.Session.md)|

## See also


[UserProperties Object Members](overview/Outlook.md)
[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
