---
title: DistListItem.RemoveMembers method (Outlook)
keywords: vbaol11.chm1155
f1_keywords:
- vbaol11.chm1155
ms.prod: outlook
api_name:
- Outlook.DistListItem.RemoveMembers
ms.assetid: 7212e075-9982-57c8-ac22-a62d3e5b3d2c
ms.date: 06/08/2017
localization_priority: Normal
---


# DistListItem.RemoveMembers method (Outlook)

Removes members from a distribution list.


## Syntax

_expression_. `RemoveMembers`( `_Recipients_` )

_expression_ A variable that represents a [DistListItem](Outlook.DistListItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Recipients_|Required| **[Recipients](Outlook.Recipients.md)**|The members to be removed from the distribution list.|

## Example

The following Microsoft Visual Basic for Applications (VBA) example removes two members from the distribution list called Group List. The **RemoveMembers** method will fail if the specified recipients are not valid. Before running the example, create or make sure a distribution list called 'Group List' exists in your default Contacts folder.


```vb
Sub RemoveRecs() 
 
 'Remove a recipient from the list and displays new list. 
 
 Dim objDstList As Outlook.DistListItem 
 
 Dim objName As Outlook.NameSpace 
 
 Dim objRcpnt As Outlook.Recipient 
 
 Dim objRcpnt2 As Outlook.Recipient 
 
 Dim objMail As Outlook.MailItem 
 
 Dim objRcpnts As Outlook.Recipients 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objDstList = objName.GetDefaultFolder(olFolderContacts).Items("Group List") 
 
 Set objMail = Application.CreateItem(olMailItem) 
 
 Set objRcpnts = objMail.Recipients 
 
 Set objRcpnt = objRcpnts.Add(Name:="someone@example.com") 
 
 Set objRcpnt2 = objRcpnts.Add(Name:="someone@example.org") 
 
 objRcpnts.ResolveAll 
 
 objDstList.RemoveMembers objRcpnts 
 
 objDstList.Display 
 
 objDstList.Body = "Last Modified: " & Now 
 
End Sub
```


## See also


[DistListItem Object](Outlook.DistListItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]