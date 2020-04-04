---
title: MeetingItem.CustomPropertyChange event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MeetingItem.CustomPropertyChange
ms.assetid: b3d05c13-4b5d-032b-49bb-18c4f4a626b5
ms.date: 06/08/2017
localization_priority: Normal
---


# MeetingItem.CustomPropertyChange event (Outlook)

Occurs when a custom property of an item (which is an instance of the parent object) is changed. 


## Syntax

_expression_. `CustomPropertyChange`( `_Name_` )

_expression_ A variable that represents a [MeetingItem](Outlook.MeetingItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the custom property that was changed.|

## Remarks

The property name is passed to the procedure so that you can determine which custom property changed.


## Example

This Microsoft Visual Basic Scripting Edition (VBScript) example uses the  **CustomPropertyChange** event to enable a control when a **Boolean** field is set to **True**.

For this example, create two custom fields on the second page of a form. The first, a **Boolean** field, is named "RespondBy". The second field is named "DateToRespond".




```vb
Sub Item_CustomPropertyChange(ByVal myPropName) 
 Select Case myPropName 
 Case "RespondBy" 
 Set myPages = Item.GetInspector.ModifiedFormPages 
 Set myCtrl = myPages("P.2").Controls("DateToRespond") 
 If Item.UserProperties("RespondBy").Value Then 
 myCtrl.Enabled = True 
 myCtrl.Backcolor = 65535 'Yellow 
 Else 
 myCtrl.Enabled = False 
 myCtrl.Backcolor = 0 'Black 
 End If 
 Case Else 
 End Select 
End Sub
```


## See also


[MeetingItem Object](Outlook.MeetingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]