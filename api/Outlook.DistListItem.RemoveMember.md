---
title: DistListItem.RemoveMember method (Outlook)
keywords: vbaol11.chm1160
f1_keywords:
- vbaol11.chm1160
ms.prod: outlook
api_name:
- Outlook.DistListItem.RemoveMember
ms.assetid: 3c0984f9-69b9-42e1-a9c2-75c60c0d0e3a
ms.date: 06/08/2017
localization_priority: Normal
---


# DistListItem.RemoveMember method (Outlook)

Removes an individual member from a given distribution list.


## Syntax

_expression_. `RemoveMember`( `_Recipient_` )

_expression_ A variable that represents a [DistListItem](Outlook.DistListItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Recipient_|Required| **[Recipient](Outlook.Recipient.md)**|The  **Recipient** to be removed from the distribution list.|

## Example

The following Microsoft Visual Basic for Applications (VBA) example removes a member from the distribution list called Group List. The  **RemoveMember** method will fail if the specified recipient is not valid. Before running the example, create or make sure a distribution list called 'Group List' exists in your default Contacts folder.


```vb
Sub RemoveRec() 
 
 'Remove a recipient from the list, and displays new list. 
 
 
 
 Dim objDstList As Outlook.DistListItem 
 
 Dim objName As Outlook.NameSpace 
 
 Dim objRcpnt As Outlook.Recipient 
 
 Dim objMail As Outlook.MailItem 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objDstList = objName.GetDefaultFolder(olFolderContacts).Items("Group List") 
 
 Set objMail = Application.CreateItem(olMailItem) 
 
 Set objRcpnt = objMail.Recipients.Add(Name:="someone@example.com") 
 
 objRcpnt.Resolve 
 
 objDstList.RemoveMember Recipient:=objRcpnt 
 
 objDstList.Display 
 
 objDstList.Body = "Last Modified: " & Now 
 
End Sub
```


## See also


[DistListItem Object](Outlook.DistListItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]