---
title: UserProperty object (Outlook)
keywords: vbaol11.chm212
f1_keywords:
- vbaol11.chm212
ms.prod: outlook
api_name:
- Outlook.UserProperty
ms.assetid: c94f642f-4368-d775-a79f-ce6c39bfe1fd
ms.date: 06/08/2017
localization_priority: Normal
---


# UserProperty object (Outlook)

Represents a custom property of an Outlook item.


## Remarks

Use  **[UserProperties](Outlook.MailItem.UserProperties.md)** (_index_), where _index_ is a name or index number, to return a single **UserProperty** object.

Use the  **[Add](Outlook.UserProperties.Add.md)** method to create a new **UserProperty** for an item and add it to the **[UserProperties](Outlook.UserProperties.md)** object. The **Add** method allows you to specify a name and type for the new property.




> [!NOTE] 
> When you create a custom property, a field is added in the folder that contains the item (using the same name as the property). That field can be used as a column in folder views.


## Example

The following example adds a custom text property named MyPropName.


```vb
Set myProp = myItem.UserProperties.Add("MyPropName", olText)
```


## Methods



|Name|
|:-----|
|[Delete](Outlook.UserProperty.Delete.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.UserProperty.Application.md)|
|[Class](Outlook.UserProperty.Class.md)|
|[Formula](Outlook.UserProperty.Formula.md)|
|[Name](Outlook.UserProperty.Name.md)|
|[Parent](Outlook.UserProperty.Parent.md)|
|[Session](Outlook.UserProperty.Session.md)|
|[Type](Outlook.UserProperty.Type.md)|
|[ValidationFormula](Outlook.UserProperty.ValidationFormula.md)|
|[ValidationText](Outlook.UserProperty.ValidationText.md)|
|[Value](Outlook.UserProperty.Value.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)
[UserProperty Object Members](overview/Outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]