---
title: PropertyAccessor object (Outlook)
keywords: vbaol11.chm3157
f1_keywords:
- vbaol11.chm3157
ms.prod: outlook
api_name:
- Outlook.PropertyAccessor
ms.assetid: 2fc91e13-703c-3ec9-9066-ffee7144306c
ms.date: 06/08/2017
localization_priority: Normal
---


# PropertyAccessor object (Outlook)

Provides the ability to create, get, set, and delete properties on objects.


## Remarks

Use the  **PropertyAccessor** object to get and set item-level properties that are not explicitly exposed in the Outlook object model, or properties for the following non-item objects: **[AddressEntry](Outlook.AddressEntry.md)**, **[AddressList](Outlook.AddressList.md)**, **[Attachment](Outlook.Attachment.md)**, **[ExchangeDistributionList](Outlook.ExchangeDistributionList.md)**, **[ExchangeUser](Outlook.ExchangeUser.md)**, **[Folder](Outlook.Folder.md)**, **[Recipient](Outlook.Recipient.md)**, and **[Store](Outlook.Store.md)**.

To get or set multiple custom properties, use the  **PropertyAccessor** object instead of the **[UserProperties](Outlook.UserProperties.md)** object for better performance.

For more information on using the  **PropertyAccessor** object, see [Properties Overview](../outlook/How-to/Navigation/properties-overview.md).


## Example

The following code sample demonstrates how to use the  **[PropertyAccessor.GetProperty](Outlook.PropertyAccessor.GetProperty.md)** method to read a MAPI property that belongs to a **[MailItem](Outlook.MailItem.md)** but that is not exposed in the Outlook object model, **PR_TRANSPORT_MESSAGE_HEADERS**.


```vb
Sub DemoPropertyAccessorGetProperty() 
 
 Dim PropName, Header As String 
 
 Dim oMail As Object 
 
 Dim oPA As Outlook.PropertyAccessor 
 
 'Get first item in the inbox 
 
 Set oMail = _ 
 
 Application.Session.GetDefaultFolder(olFolderInbox).Items(1) 
 
 'PR_TRANSPORT_MESSAGE_HEADERS 
 
 PropName = "http://schemas.microsoft.com/mapi/proptag/0x007D001E" 
 
 'Obtain an instance of PropertyAccessor class 
 
 Set oPA = oMail.PropertyAccessor 
 
 'Call GetProperty 
 
 Header = oPA.GetProperty(PropName) 
 
 Debug.Print (Header) 
 
End Sub
```

The next code sample demonstrates how the  **[PropertyAccessor.SetProperties](Outlook.PropertyAccessor.SetProperties.md)** method sets the values of multiple properties. If a property does not exist, then **SetProperties** will create the property as long as the parent object supports the creation of those properties. If the object supports an explicit **Save** operation, then the properties are saved to the object when the explicit **Save** operation is called. If the object does not support an explicit **Save** operation, then the properties are saved to the object when **SetProperties** is called.




```vb
Sub DemoPropertyAccessorSetProperties() 
 
 Dim PropNames(), myValues() As Variant 
 
 Dim arrErrors As Variant 
 
 Dim prop1, prop2, prop3, prop4 As String 
 
 Dim i As Integer 
 
 Dim oMail As Outlook.MailItem 
 
 Dim oPA As Outlook.PropertyAccessor 
 
 'Get first item in the inbox 
 
 Set oMail = _ 
 
 Application.Session.GetDefaultFolder(olFolderInbox).Items(1) 
 
 'Names for properties using the MAPI string namespace 
 
 prop1 = "http://schemas.microsoft.com/mapi/string/" & _ 
 
 "{FFF40745-D92F-4C11-9E14-92701F001EB3}/mylongprop" 
 
 prop2 = "http://schemas.microsoft.com/mapi/string/" & _ 
 
 "{FFF40745-D92F-4C11-9E14-92701F001EB3}/mystringprop" 
 
 prop3 = "http://schemas.microsoft.com/mapi/string/" & _ 
 
 "{FFF40745-D92F-4C11-9E14-92701F001EB3}/mydateprop" 
 
 prop4 = "http://schemas.microsoft.com/mapi/string/" & _ 
 
 "{FFF40745-D92F-4C11-9E14-92701F001EB3}/myboolprop" 
 
 PropNames = Array(prop1, prop2, prop3, prop4) 
 
 myValues = Array(1020, "111-222-Kudo", Now(), False) 
 
 'Set values with SetProperties call 
 
 'If the properties do not exist, then SetProperties 
 
 'adds the properties to the object when saved. 
 
 'The type of the property is the type of the element 
 
 'passed in myValues array. 
 
 Set oPA = oMail.PropertyAccessor 
 
 arrErrors = oPA.SetProperties(PropNames, myValues) 
 
 If Not (IsEmpty(arrErrors)) Then 
 
 'Examine the arrErrors array to determine if any 
 
 'elements contain errors 
 
 For i = LBound(arrErrors) To UBound(arrErrors) 
 
 'Examine the type of the element 
 
 If IsError(arrErrors(i)) Then 
 
 Debug.Print (CVErr(arrErrors(i))) 
 
 End If 
 
 Next 
 
 End If 
 
 'Save the item 
 
 oMail.Save 
 
End Sub
```


## Methods



|Name|
|:-----|
|[BinaryToString](Outlook.PropertyAccessor.BinaryToString.md)|
|[DeleteProperties](Outlook.PropertyAccessor.DeleteProperties.md)|
|[DeleteProperty](Outlook.PropertyAccessor.DeleteProperty.md)|
|[GetProperties](Outlook.PropertyAccessor.GetProperties.md)|
|[GetProperty](Outlook.PropertyAccessor.GetProperty.md)|
|[LocalTimeToUTC](Outlook.PropertyAccessor.LocalTimeToUTC.md)|
|[SetProperties](Outlook.PropertyAccessor.SetProperties.md)|
|[SetProperty](Outlook.PropertyAccessor.SetProperty.md)|
|[StringToBinary](Outlook.PropertyAccessor.StringToBinary.md)|
|[UTCToLocalTime](Outlook.PropertyAccessor.UTCToLocalTime.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.PropertyAccessor.Application.md)|
|[Class](Outlook.PropertyAccessor.Class.md)|
|[Parent](Outlook.PropertyAccessor.Parent.md)|
|[Session](Outlook.PropertyAccessor.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)
[PropertyAccessor Object Members](overview/Outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
