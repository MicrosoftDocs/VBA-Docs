---
title: ItemProperties.Item method (Outlook)
keywords: vbaol11.chm536
f1_keywords:
- vbaol11.chm536
ms.prod: outlook
api_name:
- Outlook.ItemProperties.Item
ms.assetid: 51bb7900-d3fc-650d-d43b-0da14e13ca5a
ms.date: 06/08/2017
localization_priority: Normal
---


# ItemProperties.Item method (Outlook)

Returns an **[ItemProperty](Outlook.ItemProperty.md)** object from the collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents an [ItemProperties](Outlook.ItemProperties.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|Either the zero-based index number of the object, or a value used to match the default property of an object in the collection.|

## Return value

An **ItemProperty** object that represents the specified object.


## Example

The following code sample in Microsoft Visual Basic for Applications (VBA) assumes that you have opened a mail item in an Inspector. It shows how to loop from zero (0) to the total number of properties associated with the item minus one to display the name of each property.


```vb
Sub EnumerateItemProperties() 
 
 Dim oM As Outlook.MailItem 
 
 Dim i As Integer 
 
 Set oM = Application.ActiveInspector.CurrentItem 
 
 If Not (oM Is Nothing) Then 
 
 For i = 0 To oM.ItemProperties.count - 1 
 
 Debug.Print oM.ItemProperties(i).name 
 
 Next 
 
 End If 
 
End Sub
```


## See also


[ItemProperties Object](Outlook.ItemProperties.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]