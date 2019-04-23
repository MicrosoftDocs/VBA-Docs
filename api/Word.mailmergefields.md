---
title: MailMergeFields object (Word)
ms.prod: word
ms.assetid: 9d2dfd45-c52b-500e-15bf-1e678e6c1e92
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMergeFields object (Word)

A collection of  **[MailMergeField](Word.MailMergeField.md)** objects that represent the mail merge related fields in a document.


## Remarks

Use the  **Fields** property to return the **MailMergeFields** collection. The following example adds an ASK field after the last mail merge field in the active document.


```vb
Set myMMFields = ActiveDocument.MailMerge.Fields 
myMMFields(myMMFields.Count).Select 
Selection.MoveRight Unit:=wdWord, Count:=1, Extend:=wdMove 
ActiveDocument.MailMerge.Fields.AddAsk Range:=Selection.Range, _ 
 Name:="Name", Prompt:="Type your name", AskOnce:=True
```

Use the  **Add** method to add a merge field to the **MailMergeFields** collection. The following example replaces the selection with a **MiddleInitial** merge field.




```vb
ActiveDocument.MailMerge.Fields.Add Range:=Selection.Range, _ 
 Name:="MiddleInitial"
```

Use  **Fields** (Index), where Index is the index number, to return a single **MailMergeField** object. The following example displays the field code of the first mail merge field in the active document.




```vb
MsgBox ActiveDocument.MailMerge.Fields(1).Code
```

The  **MailMergeFields** collection has additional methods, such as **AddAsk** and **AddFillIn**, for adding fields related to a mail merge operation.


## Methods



|Name|
|:-----|
|[Add](Word.MailMergeFields.Add.md)|
|[AddAsk](Word.MailMergeFields.AddAsk.md)|
|[AddFillIn](Word.MailMergeFields.AddFillIn.md)|
|[AddIf](Word.MailMergeFields.AddIf.md)|
|[AddMergeRec](Word.MailMergeFields.AddMergeRec.md)|
|[AddMergeSeq](Word.MailMergeFields.AddMergeSeq.md)|
|[AddNext](Word.MailMergeFields.AddNext.md)|
|[AddNextIf](Word.MailMergeFields.AddNextIf.md)|
|[AddSet](Word.MailMergeFields.AddSet.md)|
|[AddSkipIf](Word.MailMergeFields.AddSkipIf.md)|
|[Item](Word.MailMergeFields.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Word.MailMergeFields.Application.md)|
|[Count](Word.MailMergeFields.Count.md)|
|[Creator](Word.MailMergeFields.Creator.md)|
|[Parent](Word.MailMergeFields.Parent.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]