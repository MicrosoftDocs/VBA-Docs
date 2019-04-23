---
title: DistListItem.ShowCategoriesDialog method (Outlook)
keywords: vbaol11.chm1158
f1_keywords:
- vbaol11.chm1158
ms.prod: outlook
api_name:
- Outlook.DistListItem.ShowCategoriesDialog
ms.assetid: 47cb9ecd-6d2c-53d5-e083-09935d91a510
ms.date: 06/08/2017
localization_priority: Normal
---


# DistListItem.ShowCategoriesDialog method (Outlook)

Displays the  **Show Categories** dialog box, which allows you to select categories that correspond to the subject of the item.


## Syntax

_expression_. `ShowCategoriesDialog`

_expression_ A variable that represents a [DistListItem](Outlook.DistListItem.md) object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates a new distribution list item, displays the item on the screen, and opens up the  **Show Categories** dialog box.


```vb
Sub Appointment() 
 
 'Creates an distribution list item to access ShowCategoriesDialog 
 
 Dim olDistListItem As Outlook.DistListItem 
 
 'Create distribution list item 
 
 Set olDistListItem = Application.CreateItem(olDistributionListItem) 
 
 
 
 'Display the item 
 
 olDistListItem.Display 
 
 'Display the Show categories dialog 
 
 olDistListItem.ShowCategoriesDialog 
 
End Sub
```


## See also


[DistListItem Object](Outlook.DistListItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]