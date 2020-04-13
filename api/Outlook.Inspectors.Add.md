---
title: Inspectors.Add method (Outlook)
keywords: vbaol11.chm139
f1_keywords:
- vbaol11.chm139
ms.prod: outlook
api_name:
- Outlook.Inspectors.Add
ms.assetid: f83a1cac-8103-003b-4389-d4f596e78aaa
ms.date: 06/08/2017
localization_priority: Normal
---


# Inspectors.Add method (Outlook)

Creates a new inspector window.


## Syntax

_expression_.**Add** `_Item_`

_expression_ A variable that represents an [Inspectors](Outlook.Inspectors.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The item to display in the inspector window when it is created.|

## Return value

An **[Inspector](Outlook.Inspector.md)** object that represents a new inspector window.


## Remarks

This method is essentially identical to the  **GetInspector** property of an Outlook item, such as **[MailItem](Outlook.MailItem.md)**.


## Example

This Microsoft Visual Basic for Applications (VBA) example prompts the user for a company name, uses the  **[Restrict](Outlook.Items.Restrict.md)** method to locate all contact items in the Contacts folder with that name, and displays each one.


```vb
Sub DisplayMyContacts() 
 
 Dim myFolder As Folder 
 
 Dim myItems As Items 
 
 Dim myRestrictItems As Items 
 
 Dim answer As String 
 
 Dim filter As String 
 
 Dim myInspector As Inspector 
 
 Dim x As Integer 
 
 
 
 answer = InputBox("Enter the company name") 
 
 Set myFolder = Application.GetNamespace("MAPI") _ 
 
 .GetDefaultFolder(olFolderContacts) 
 
 filter = "[MessageClass] = 'IPM.Contact' AND [CompanyName] = '" & answer & "'" 
 
 
 
 Set myItems = myFolder.Items 
 
 Set myRestrictItems = myItems.Restrict(filter) 
 
 For x = 1 To myRestrictItems.Count 
 
 Set myInspector = Application.Inspectors.Add(myRestrictItems.Item(x)) 
 
 myInspector.Display 
 
 Next x 
 
End Sub
```


## See also


[Inspectors Object](Outlook.Inspectors.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]