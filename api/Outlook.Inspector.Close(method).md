---
title: Inspector.Close method (Outlook)
keywords: vbaol11.chm2965
f1_keywords:
- vbaol11.chm2965
ms.prod: outlook
api_name:
- Outlook.Inspector.Close
ms.assetid: de821cf4-72f8-ba62-3d8d-96548db0b4a0
ms.date: 06/08/2017
localization_priority: Normal
---


# Inspector.Close method (Outlook)

Closes the  **[Inspector](Outlook.Inspector.md)** and optionally saves changes to the displayed Outlook item.


## Syntax

_expression_.**Close** (_SaveMode_)

_expression_ A variable that represents an [Inspector](Outlook.Inspector.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SaveMode_|Required| **[OlInspectorClose](Outlook.OlInspectorClose.md)**|The close behavior. If the item displayed within the inspector has not been changed, this argument has no effect.|

## Remarks


> [!NOTE] 
> Do not use this method from within the [Inspector.Activate event (Outlook)](Outlook.Inspector.Activate(even).md) event handler.


## Example

This Visual Basic for Applications (VBA) example saves and closes the item displayed in the active inspector without prompting the user. To run this example, you need to have an item displayed in an inspector window.


```vb
Sub CloseItem() 
 
 Dim myinspector As Outlook.Inspector 
 
 Dim myItem As Outlook.MailItem 
 
 
 
 Set myinspector = Application.ActiveInspector 
 
 Set myItem = myinspector.CurrentItem 
 
 myItem.Close olSave 
 
End Sub
```


## See also


[Inspector Object](Outlook.Inspector.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]