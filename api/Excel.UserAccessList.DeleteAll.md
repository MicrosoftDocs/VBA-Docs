---
title: UserAccessList.DeleteAll method (Excel)
keywords: vbaxl10.chm726076
f1_keywords:
- vbaxl10.chm726076
ms.prod: excel
api_name:
- Excel.UserAccessList.DeleteAll
ms.assetid: c162c9cf-8257-e97a-ebe8-ab1d700924ca
ms.date: 05/18/2019
localization_priority: Normal
---


# UserAccessList.DeleteAll method (Excel)

Removes all users who have access to a protected range on a worksheet.


## Syntax

_expression_.**DeleteAll**

_expression_ A variable that represents a **[UserAccessList](Excel.UserAccessList.md)** object.


## Example

In this example, Microsoft Excel removes all users that have access to the first protected range on the active worksheet. This example assumes that the worksheet is not protected.

```vb
Sub UseDeleteAll() 
 
 Dim wksSheet As Worksheet 
 
 Set wksSheet = Application.ActiveSheet 
 
 ' Remove all users with access to the first protected range. 
 wksSheet.Protection.AllowEditRanges(1).Users.DeleteAll 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]