---
title: ContactItem.LastModificationTime property (Outlook)
keywords: vbaol11.chm943
f1_keywords:
- vbaol11.chm943
ms.prod: outlook
api_name:
- Outlook.ContactItem.LastModificationTime
ms.assetid: 14962429-bbf6-a814-647a-70be1dad339d
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.LastModificationTime property (Outlook)

Returns a **Date** specifying the date and time that the Outlook item was last modified. Read-only.


## Syntax

_expression_. `LastModificationTime`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagLastModificationTime**.


## Example

This Visual Basic for Applications example uses the  **[Items.Restrict](Outlook.Items.Restrict.md)** method to apply a filter to contact items based on the item's **LastModificationTime** property. You can apply a similar approach to filter on the **LastModificationTime** property of other Outlook items.


```vb
Public Sub ContactDateCheck() 
 
 Dim myNamespace As Outlook.NameSpace 
 
 Dim myContacts As Outlook.Items 
 
 Dim myItems As Outlook.Items 
 
 Dim myItem As Object 
 
 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set myContacts = myNamespace.GetDefaultFolder(olFolderContacts).Items 
 
 Set myItems = myContacts.Restrict("[LastModificationTime] > '01/1/2003'") 
 
 For Each myItem In myItems 
 
 If (myItem.Class = olContact) Then 
 
 MsgBox myItem.FullName & ": " & myItem.LastModificationTime 
 
 End If 
 
 Next 
 
End Sub
```

The following Visual Basic for Applications example is the same as the example above, except that it demonstrates the use of a variable in the filter.




```vb
Public Sub ContactDateCheck2() 
 
 Dim myNamespace As Outlook.NameSpace 
 
 Dim myContacts As Outlook.Items 
 
 Dim myItem As Object 
 
 Dim DateStart As Date 
 
 Dim DateToCheck As String 
 
 Dim myRestrictItems As Outlook.Items 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 Set myContacts = myNameSpace.GetDefaultFolder(olFolderContacts).Items 
 
 DateStart = #01/1/2003# 
 
 DateToCheck = "[LastModificationTime] >= """ & DateStart & """" 
 
 Set myRestrictItems = myContacts.Restrict(DateToCheck) 
 
 For Each myItem In myRestrictItems 
 
 If (myItem.Class = olContact) Then 
 
 MsgBox myItem.FullName & ": " & myItem.LastModificationTime 
 
 End If 
 
 Next 
 
End Sub
```


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]