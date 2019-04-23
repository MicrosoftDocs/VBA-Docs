---
title: Items.ItemChange event (Outlook)
keywords: vbaol11.chm315
f1_keywords:
- vbaol11.chm315
ms.prod: outlook
api_name:
- Outlook.Items.ItemChange
ms.assetid: 6478357e-2a5a-300a-24e6-c125f8c81edd
ms.date: 06/08/2017
localization_priority: Normal
---


# Items.ItemChange event (Outlook)

Occurs when an item in the specified collection is changed. This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Syntax

_expression_. `ItemChange`( `_Item_` )

_expression_ A variable that represents an [Items](Outlook.Items.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The item that was changed.|

## Example

This example uses the  **[Start](Outlook.AppointmentItem.Start.md)** property of the **[AppointmentItem](Outlook.AppointmentItem.md)** object to determine if the appointment starts after normal business hours. If it does, and if the **[Sensitivity](Outlook.AppointmentItem.Sensitivity.md)** property of the **AppointmentItem** object is not already set to **olPrivate**, the example offers to mark the appointment as private.


```vb
Public WithEvents myOlItems As Outlook.Items 
 
 
 
Public Sub Initialize_handler() 
 
 Set myOlItems = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar).Items 
 
End Sub 
 
 
 
Private Sub myOlItems_ItemChange(ByVal Item As Object) 
 
Dim prompt As String 
 
 If VBA.Format(Item.Start, "h") >= "17" And Item.Sensitivity <> olPrivate Then 
 
 prompt = "Appointment occurs after hours. Mark it private?" 
 
 If MsgBox(prompt, vbYesNo + vbQuestion) = vbYes Then 
 
 Item.Sensitivity = olPrivate 
 
 Item.Display 
 
 End If 
 
 End If 
 
End Sub
```


## See also


[Items Object](Outlook.Items.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]