---
title: DistListItem.DLName property (Outlook)
keywords: vbaol11.chm1148
f1_keywords:
- vbaol11.chm1148
ms.prod: outlook
api_name:
- Outlook.DistListItem.DLName
ms.assetid: 38d027b7-89f9-1659-84e0-35473b07c088
ms.date: 06/08/2017
localization_priority: Normal
---


# DistListItem.DLName property (Outlook)

Returns or sets a  **String** representing the display name of a distribution list. Read/write.


## Syntax

_expression_. `DLName`

_expression_ A variable that represents a [DistListItem](Outlook.DistListItem.md) object.


## Example

This Microsoft Visual Basic for Applications (VBA) example creates a new distribution list and then prompts the user for a name.


```vb
Sub CreateDL() 
 
 Dim myDistList As Outlook.DistListItem 
 
 
 
 Set myDistList = Application.CreateItem(olDistributionListItem) 
 
 myDistList.DLName = InputBox("Type the name of the new distribution list.") 
 
 myDistList.Save 
 
 myDistList.Display 
 
End Sub
```


## See also


[DistListItem Object](Outlook.DistListItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]