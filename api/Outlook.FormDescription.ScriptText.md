---
title: FormDescription.ScriptText property (Outlook)
keywords: vbaol11.chm197
f1_keywords:
- vbaol11.chm197
ms.prod: outlook
api_name:
- Outlook.FormDescription.ScriptText
ms.assetid: 56ea4cd6-a9f0-cd0c-a378-dab6399bd1ca
ms.date: 06/08/2017
localization_priority: Normal
---


# FormDescription.ScriptText property (Outlook)

Returns a  **String** containing all the VBScript code in the form's Script Editor. Read-only.


## Syntax

_expression_. `ScriptText`

_expression_ A variable that represents a '[FormDescription](Outlook.FormDescription.md)' object.


## Example

This Microsoft Visual Basic Scripting Edition (VBScript) example uses the  **[Open](Outlook.MailItem.Open.md)** event to access the **[HTMLBody](Outlook.MailItem.HTMLBody.md)** property of a **[MailItem](Outlook.MailItem.md)**. This sets the **[EditorType](Outlook.Inspector.EditorType.md)** property of the **MailItem** 's **[Inspector](Outlook.Inspector.md)** to **olEditorHTML**. When the **MailItem** 's **[Body](Outlook.MailItem.Body.md)** property is set, the **EditorType** property is changed to the default. For example, if the default email editor is set to RTF, the **EditorType** is set to **olEditorRTF**. If this code is placed in the Script Editor of a form in design mode, the message boxes during run time will reflect the change in the **EditorType** as the body of the form changes. The final message box uses the **Script Text** property to display all the VBScript code in the Script Editor.


```vb
Function Item_Open() 
 
 'Set the HTMLBody of the item. 
 
 Item.HTMLBody = "<HTML><H2>My HTML page.</H2><BODY>My body.</BODY></HTML>" 
 
 'Item displays HTML message. 
 
 Item.Display 
 
 'MsgBox shows EditorType is 2. 
 
 MsgBox "HTMLBody EditorType is " & Item.GetInspector.EditorType 
 
 'Access the Body and show 
 
 'the text of the Body. 
 
 MsgBox "This is the Body: " & Item.Body 
 
 'After accessing, EditorType 
 
 'is still 2. 
 
 MsgBox "After accessing, the EditorType is " & Item.GetInspector.EditorType 
 
 'Set the item's Body property. 
 
 Item.Body = "Back to default body." 
 
 'After setting, EditorType is 
 
 'now back to the default. 
 
 MsgBox "After setting, the EditorType is " & Item.GetInspector.EditorType 
 
 'Access the items's 
 
 'FormDescription object. 
 
 Set myForm = Item.FormDescription 
 
 'Display all the code 
 
 'in the Script Editor. 
 
 MsgBox myForm.ScriptText 
 
End Function
```


## See also


[FormDescription Object](Outlook.FormDescription.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]