---
title: MailItem.GetInspector property (Outlook)
keywords: vbaol11.chm1305
f1_keywords:
- vbaol11.chm1305
ms.prod: outlook
api_name:
- Outlook.MailItem.GetInspector
ms.assetid: 9ba8bdbf-1dd5-eaff-3889-33433e3cb3fa
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.GetInspector property (Outlook)

Returns an  **[Inspector](Outlook.Inspector.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

_expression_. `GetInspector`

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Remarks

This property is useful for returning an  **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](Outlook.Application.ActiveInspector.md)** method and setting the **[Inspector.CurrentItem](Outlook.Inspector.CurrentItem.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## Example

This Visual Basic for Applications (VBA) example shows a function  `InsertBodyTextInWordEditor` that creates a mail item, assigns it a title and adds text for the body. The function sets the **[Subject](Outlook.MailItem.Subject.md)** property to assign the title "Testing...". It then calls the **[Display](Outlook.MailItem.Display.md)** method to open the mail item in an inspector. To insert text in a Word editor as the body of the mail item, the function uses the **[Document](./Word.Document.md)** object and **[Range](./Word.Range.md)** object in the Word object model. The function uses the item's **GetInspector** property to get the existing **Inspector** object, and then uses the **[Inspector.WordEditor](Outlook.Inspector.WordEditor.md)** property to obtain a **Word.Document** object for the item. Using the **Word.Document** object, the function accesses the **Word.Range** object and inserts text into the body of the item.

Since this example accesses the Word object model, you must first add a reference to the Microsoft Word Object Library to compile the example successfully.




```vb
Sub InsertBodyTextInWordEditor() 
 Dim myItem As Outlook.MailItem 
 Dim myInspector As Outlook.Inspector 
 'You must add a reference to the Microsoft Word Object Library 
 'before this sample will compile 
 Dim wdDoc As Word.Document 
 Dim wdRange As Word.Range 
 
 On Error Resume Next 
 Set myItem = Application.CreateItem(olMailItem) 
 myItem.Subject = "Testing..." 
 myItem.Display 
 'GetInspector property returns Inspector 
 Set myInspector = myItem.GetInspector 
 'Obtain the Word.Document for the Inspector 
 Set wdDoc = myInspector.WordEditor 
 If Not (wdDoc Is Nothing) Then 
 'Use the Range object to insert text 
 Set wdRange = wdDoc.Range(0, wdDoc.Characters.Count) 
 wdRange.InsertAfter ("Hello world!") 
 End If 
End Sub
```


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
