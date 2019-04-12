---
title: Items.Sort method (Outlook)
keywords: vbaol11.chm72
f1_keywords:
- vbaol11.chm72
ms.prod: outlook
api_name:
- Outlook.Items.Sort
ms.assetid: 7cb248a2-6885-8be5-df7b-fd5683081e01
ms.date: 06/08/2017
localization_priority: Normal
---


# Items.Sort method (Outlook)

Sorts the collection of items by the specified property. The index for the collection is reset to 1 upon completion of this method.


## Syntax

_expression_.**Sort** (_Property_, _Descending_)

_expression_ A variable that represents an [Items](Outlook.Items.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Property_|Required| **String**|The name of the property by which to sort, which may be enclosed in brackets, for example, "[CompanyName]". User-defined properties that contain spaces must be enclosed in brackets. May not be a user-defined property of type keywords, and may not be a multi-valued property, such as a category. For user-defined properties, the property must exist in the  **UserDefinedProperties** collection for **[Items.Parent](Outlook.Items.Parent.md)**, which represents the **[Folder](Outlook.Folder.md)** object that contains the items.|
| _Descending_|Optional| **Variant**| **True** to sort in descending order. The default value is **False** (ascending).|

## Remarks

 **Sort** only affects the order of items in a collection. It does not affect the order of items in an explorer view.

 **Sort** cannot be used and will cause an error if the _Property_ paramater is one of the following properties:



| **Categories**| **[LastFirstSpaceOnly](Outlook.ContactItem.LastFirstSpaceOnly.md)**|
| **[Children](Outlook.ContactItem.Children.md)**| **[LastFirstSpaceOnlyCompany](Outlook.ContactItem.LastFirstSpaceOnlyCompany.md)**|
| **Class**| **[MemberCount](Outlook.DistListItem.MemberCount.md)**|
| **[CompanyLastFirstNoSpace](Outlook.ContactItem.CompanyLastFirstNoSpace.md)**| **[NetMeetingAlias](Outlook.ContactItem.NetMeetingAlias.md)**|
| **[CompanyLastFirstSpaceOnly](Outlook.ContactItem.CompanyLastFirstSpaceOnly.md)**| **[RecurrenceState](Outlook.AppointmentItem.RecurrenceState.md)**|
| **[DLName](Outlook.DistListItem.DLName.md)**| **[ResponseState](Outlook.TaskItem.ResponseState.md)**|
| **[LastFirstAndSuffix](Outlook.ContactItem.LastFirstAndSuffix.md)**| **Saved**|
| **[LastFirstNoSpace](Outlook.ContactItem.LastFirstNoSpace.md)**| **Sent**|
| **[LastFirstNoSpaceCompany](Outlook.ContactItem.LastFirstNoSpaceCompany.md)**||

## Example

The following Visual Basic for Applications (VBA) example uses the  **Sort** method to sort the **[Items](Outlook.Items.md)** collection for the default **Tasks** folder by the "DueDate" property and displays the due dates each in turn.


```vb
Sub SortByDueDate() 
 Dim myNameSpace As Outlook.NameSpace 
 Dim myFolder As Outlook.Folder 
 Dim myItem As Outlook.TaskItem 
 Dim myItems As Outlook.Items 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderTasks) 
 Set myItems = myFolder.Items 
 myItems.Sort "[DueDate]", False 
 For Each myItem In myItems 
 MsgBox myItem.Subject & "-- " & myItem.DueDate 
 Next myItem 
End Sub
```


## See also


[Items Object](Outlook.Items.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
