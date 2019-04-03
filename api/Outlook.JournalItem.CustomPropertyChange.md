---
title: JournalItem.CustomPropertyChange event (Outlook)
ms.prod: outlook
api_name:
- Outlook.JournalItem.CustomPropertyChange
ms.assetid: bdaad359-bc21-c8a9-c934-7acf92d836ae
ms.date: 06/08/2017
localization_priority: Normal
---


# JournalItem.CustomPropertyChange event (Outlook)

Occurs when a custom property of an item (which is an instance of the parent object) is changed. 


## Syntax

_expression_. `CustomPropertyChange`( `_Name_` )

_expression_ A variable that represents a [JournalItem](Outlook.JournalItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the custom property that was changed.|

## Remarks

The property name is passed to the procedure so that you can determine which custom property changed.


## Example

This Microsoft Visual Basic Scripting Edition (VBScript) example uses the  **CustomPropertyChange** event to enable a control when a **Boolean** field is set to **True**.

For this example, create two custom fields on the second page of a form. The first, a  **Boolean** field, is named "RespondBy". The second field is named "DateToRespond".




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


[JournalItem Object](Outlook.JournalItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]