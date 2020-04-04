---
title: Inspector.EditorType property (Outlook)
keywords: vbaol11.chm2963
f1_keywords:
- vbaol11.chm2963
ms.prod: outlook
api_name:
- Outlook.Inspector.EditorType
ms.assetid: b19e552b-1e8a-8915-f793-396860910f40
ms.date: 06/08/2017
localization_priority: Normal
---


# Inspector.EditorType property (Outlook)

Returns an **[OlEditorType](Outlook.OlEditorType.md)** constant indicating the type of editor. Read-only.


## Syntax

_expression_. `EditorType`

_expression_ A variable that represents an [Inspector](Outlook.Inspector.md) object.


## Remarks

Since Microsoft Office Outlook 2007, the  **EditorType** property always returns **olEditorWord**.


## Example

This Microsoft Visual Basic Scripting Edition (VBScript) example uses the  **[Open](Outlook.MailItem.Open.md)** event to access the **[HTMLBody](Outlook.MailItem.HTMLBody.md)** property of an item. This sets the **[EditorType](Outlook.Inspector.EditorType.md)** property of the item's **[Inspector](Outlook.Inspector.md)** to **olEditorHTML**. If this code is placed in the Script Editor of a form in design mode, the message boxes during run time will reflect the change in the **EditorType** as the body of the form changes. The final message box utilizes the **[ScriptText](Outlook.FormDescription.ScriptText.md)** property to display all the VBScript code in the Script Editor.


```vb
Function Item_Open() 
 'Set the HTMLBody of the item. 
 Item.HTMLBody = "<HTML><H2>My HTML page.</H2><BODY>My body.</BODY></HTML>" 
 'Item displays HTML message. 
 Item.Display 
 'MsgBox shows EditorType is 2 which represents the HTML editor type 
 MsgBox "HTMLBody EditorType is " & Item.GetInspector.EditorType 
 'Access the Body and show 
 'the text of the Body. 
 MsgBox "This is the Body: " & Item.Body 
 'After accessing, EditorType 
 'is still 2. 
 MsgBox "After accessing, the EditorType is " & Item.GetInspector.EditorType 
 'Set the item's Body property. 
 Item.Body = "Back to default body." 
 'After setting the Body, EditorType is 
 'still the same. 
 MsgBox "After setting, the EditorType is " & Item.GetInspector.EditorType 
End Function
```


## See also


[Inspector Object](Outlook.Inspector.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]