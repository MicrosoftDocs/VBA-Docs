---
title: Best Practices for Getting and Setting Properties
ms.prod: outlook
ms.assetid: ec087bf8-cfac-9b20-3cb2-3bd308c5c63d
ms.date: 06/08/2017
localization_priority: Normal
---


# Best Practices for Getting and Setting Properties

Keep in mind the following best practices recommendations for getting and setting values for properties:


- Reference a property directly off the parent object to get and set explicit built-in properties of item objects, for example,  **[MailItem.Subject](../../../api/Outlook.MailItem.Subject.md)**.
    
- Use  **[ItemProperties](../../../api/Outlook.ItemProperties.md)** and **[ItemProperty](../../../api/Outlook.ItemProperty.md)** to enumerate explicit built-in properties and custom properties, and get and set custom properties for items (except for **[DocumentItem](../../../api/Outlook.DocumentItem.md)** object).
    
- Use  **[UserProperties](../../../api/Outlook.UserProperties.md)** and **[UserProperty](../../../api/Outlook.UserProperty.md)** to enumerate, get and set custom properties for items (except for the **DocumentItem** object).
    
- Use  **[PropertyAccessor](../../../api/Outlook.PropertyAccessor.md)** to get and set custom properties for the **DocumentItem** object, built-in item-level properties that are not exposed in the Outlook object model, or properties for the following objects: **[AddressEntry](../../../api/Outlook.AddressEntry.md)**,  **[AddressList](../../../api/Outlook.AddressList.md)**,  **[Attachment](../../../api/Outlook.Attachment.md)**,  **[ExchangeDistributionList](../../../api/Outlook.ExchangeDistributionList.md)**,  **[ExchangeUser](../../../api/Outlook.ExchangeUser.md)**,  **[Folder](../../../api/Outlook.Folder.md)**,  **[Recipient](../../../api/Outlook.Recipient.md)**, and  **[Store](../../../api/Outlook.Store.md)**.
    
- To get or set multiple custom properties, use the  **PropertyAccessor** object instead of the **UserProperties** object for better performance.
    
- To create or access custom properties, use the MAPI string namespace for convenience over the MAPI proptag or id namespace. Use the GUID of your add-in as the namespace GUID.
    
- When referencing properties by namespaces, be aware that such references are case-sensitive. For example, while  **urn:schemas:contacts:givenName** is a valid namespace, **urn:schemas:contacts:givenname** is not.
    
- To get or set multiple properties, use  **[PropertyAccessor.GetProperties](../../../api/Outlook.PropertyAccessor.GetProperties.md)** and **[PropertyAccessor.SetProperties](../../../api/Outlook.PropertyAccessor.SetProperties.md)**, as opposed to repeated  **[PropertyAccessor.GetProperty](../../../api/Outlook.PropertyAccessor.GetProperty.md)** and **[PropertyAccessor.SetProperty](../../../api/Outlook.PropertyAccessor.SetProperty.md)**, for better performance.
    
- To have the  **CustomPropertyChange** event fire when the value of an item-level custom property changes, the custom property must be in the item's **UserProperties** collection. An item-level property added implicitly by **SetProperty** or **SetProperties** does not automatically become part of the item's **UserProperties** collection. An explicit **[UserProperties.Add](../../../api/Outlook.UserProperties.Add.md)** is required to include it.
    
- To set for the first time a property created by the  **UserProperties.Add** method, use the **[UserProperty.Value](../../../api/Outlook.UserProperty.Value.md)** property instead of the **SetProperties** and **SetProperty** methods of the **PropertyAccessor** object.
    

This section describes the best practices for saving properties on an object:


- For item objects, calling the item's  **Save** method to save the item to the current folder also saves its properties on the item.
    
- For non-item-level objects that do not have a  **Save** method (**AddressList**,  **Folder**,  **Recipient**, and  **Store**), calling  **[PropertyAccessor.DeleteProperty](../../../api/Outlook.PropertyAccessor.DeleteProperty.md)**,  **[PropertyAccessor.DeleteProperties](../../../api/Outlook.PropertyAccessor.DeleteProperties.md)**,  **SetProperty**, or  **SetProperties** will implicitly save the properties on the object.
    
This section describes the best practices for keeping type conversion simple when using the  **PropertyAccessor** to get and set properties. For definitions of MAPI property types such as **PT_SYSTIME**, see  [Property Types](../../../api/overview/Outlook.md).

- Although most Outlook date-time values are stored in Coordinated Universal Time (UTC) format, there is no guarantee that all properties of the MAPI type  **PT_SYSTIME** will always return UTC. Getting a **PT_SYSTIME** property will return a **VT_DATE** value. When setting a **PT_SYSTIME** property, ensure that you are setting the property as a UTC value rather than a local date-time value. The **GetProperty**,  **SetProperty**,  **GetProperties**, and  **SetProperties** methods do not perform time zone conversion. Use the helper methods **[PropertyAccessor.LocalTimeToUTC](../../../api/Outlook.PropertyAccessor.LocalTimeToUTC.md)** and **[PropertyAccessor.UTCToLocalTime](../../../api/Outlook.PropertyAccessor.UTCToLocalTime.md)** to perform explicit time zone conversion.
    
- A multi-valued property (Microsoft Visual Basic type  **VT_ARRAY**) is stored as a two-dimensional array that contains the same number of elements as are there are values in the property. Getting a multi-valued property will return a  **VT_ARRAY** value. When setting a multi-valued property, pass a two-dimensional array (**VT_ARRAY**) with one element for each value that you want to set for the property.
    
- A binary property (MAPI type  **PT_BINARY**) is stored as an array of bytes rather than a string. Getting a binary property will return a value of type  **VT_ARRAY**. The  **GetProperty**,  **SetProperty**,  **GetProperties**, and  **SetProperties** methods do not perform any conversion between binary array and string. Use the helper methods **[PropertyAccessor.BinaryToString](../../../api/Outlook.PropertyAccessor.BinaryToString.md)** and **[PropertyAccessor.StringToBinary](../../../api/Outlook.PropertyAccessor.StringToBinary.md)** to explicitly perform any conversion.
    
- Certain MAPI property types, such as  **PT_OBJECT**, are not supported by the  **PropertyAccessor**. Attempting to get or set such properties will result in a "property operation not supported" error.
    
- When getting or setting a property using a reference in the MAPI proptag namespace, make sure that the type specified in the proptag matches the underlying type of the property. Except for the case of a  **PT_STRING8** property where you can specify either a type of 001E or 001F in the proptag to get or set the property as a **VT_BSTR**, getting or setting a property does not involve any type coercion and an error will be returned if there is a type mismatch.
    
- When setting a property, it may be less restrictive to use a property reference in the MAPI string namespace than one in the MAPI proptag namespace. Specifying the property in the MAPI string namespace does not strictly require the value to match the underlying type of the property. For example, you can pass a string value like  **VT_BSTR** to set a date-time property such as **PT_SYSTIME**, and the type of the property becomes the type of the value, which is  **VT_BSTR**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]