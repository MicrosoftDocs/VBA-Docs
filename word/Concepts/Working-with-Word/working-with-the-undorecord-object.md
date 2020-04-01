---
title: Working with the UndoRecord Object
ms.prod: word
ms.assetid: e9df1047-5a1a-91da-3673-7e64b668552d
ms.date: 06/08/2017
localization_priority: Normal
---


# Working with the UndoRecord Object

As end users work in Word, the actions they take that can be undone are recorded in a queue—the "undo stack"—so that if the user decides to reverse the action, Word can recall the action taken and undo it. The undo stack is visible through the Word user interface on the ribbon through the Undo button, where each action is listed as a discrete record in the stack. Automation using the Word object model also writes these "undo records" to the undo stack, but working with undo stack using can be problematic.

Developers sometimes write code that performs one logical task using a series of actions, where each action is recorded as a separate entry on the undo stack. However, even though the end user can undo each action separately, sometimes it is logical to undo all the steps of the task as a unit and not individually.

For example, a developer might write code that performs the task of inserting document metadata into the footer of a Word document as shown in the following code example.




```vb
Sub AddDocMetadata() 
Dim rngFooter As Range 
 
    Set rngFooter = ActiveDocument.Sections(1) _ 
        .Footers(wdHeaderFooterPrimary).Range 
         
    With rngFooter 
        .Delete 
        .Fields.Add Range:=rngFooter, Type:=wdFieldFileName, Text:="\p" 
        .InsertAfter Text:=vbTab & vbTab 
        .Collapse Direction:=wdCollapseStart 
        .Fields.Add Range:=rngFooter, Type:=wdFieldAuthor 
    End With     
End Sub 

```

The resulting undo records appear on the undo stack for the actions performed by the code.

![Undo button expanded with several entries.](../../../images/wd14_VBAConceptual_WorkingWithUndoRecord_fig01.jpg)

If the end user then needs to rollback or undo the "insert metadata" task, each action associated with the task must be undone individually (by pressing CTRL+Z or clicking the dropdown arrow on the **Undo** button and then selecting the actions to undo). In this situation, the potential exists for the end user to only partially undo the series of actions associated with the single task, possibly leaving the document in an undesired state.

The Word object model offers a way to avoid this problem by providing developers the ability to link a series of individual actions into a custom undo record that can then be used to undo the actions as a single unit. A custom undo record is represented in the Word object model by the [UndoRecord](../../../api/Word.UndoRecord.md) object.

## Using the UndoRecord Object

The [UndoRecord](../../../api/Word.UndoRecord.md) object is a property of the [Application](../../../api/Word.Application.md) object in Word. You create a custom undo record using the [StartCustomRecord](../../../api/Word.UndoRecord.StartCustomRecord.md) and [EndCustomRecord](../../../api/Word.UndoRecord.EndCustomRecord.md) methods of the [UndoRecord](../../../api/Word.UndoRecord.md) object. Consider the following code example which modifies the previous example to include a custom undo record.


```vb
Dim objUndo As UndoRecord 
 
Sub AddDocMetadata() 
Dim rngFooter As Range 
 
Set objUndo = Application.UndoRecord 
 
'Begin the custom undo record and provide a name for the record 
objUndo.StartCustomRecord ("Add Doc Metadata") 
     
 Set rngFooter = ActiveDocument.Sections(1) _ 
        .Footers(wdHeaderFooterPrimary).Range 
         
 With rngFooter 
        .Delete 
        .Fields.Add Range:=rngFooter, Type:=wdFieldFileName, Text:="\p" 
        .InsertAfter Text:=vbTab & vbTab 
        .Collapse Direction:=wdCollapseStart 
        .Fields.Add Range:=rngFooter, Type:=wdFieldAuthor 
 End With 
 
'End the custom undo record 
objUndo.EndCustomRecord 
     
End Sub 

```

The [StartCustomRecord](../../../api/Word.UndoRecord.StartCustomRecord.md) method begins recording the actions that are included in the custom undo record. You can give the custom undo record a name as an argument to the [StartCustomRecord](../../../api/Word.UndoRecord.StartCustomRecord.md) method, but this is optional. If you do not specify a name, Word uses the Visual Basic name of the first command executed as the name of the record. The name of the custom undo record is the string that appears in the user interface in the dropdown next to the **Undo** button.

The [EndCustomRecord](../../../api/Word.UndoRecord.EndCustomRecord.md) method stops recording the custom undo record. All the actions taken by the code between the [StartCustomRecord](../../../api/Word.UndoRecord.StartCustomRecord.md) call and the [EndCustomRecord](../../../api/Word.UndoRecord.EndCustomRecord.md) call are linked into one [UndoRecord](../../../api/Word.UndoRecord.md) object. When you run code that creates a custom undo record, as in the previous code example, the custom undo record is listed in the undo stack as a single entry.


![Undo button expanded with one entry.](../../../images/wd14_VBAConceptual_WorkingWithUndoRecord_fig02.jpg)


## UndoRecord Object Members

The following table lists other important members of the [UndoRecord](../../../api/Word.UndoRecord.md).



|**Name**|**Member Type**|**Description**|
|:-----|:-----|:-----|
|**[CustomRecordLevel](../../../api/Word.UndoRecord.CustomRecordLevel.md)**|Property|Returns a **Long** that specifies the number of custom undo action calls that are currently active. Read-only.|
|**[CustomRecordName](../../../api/Word.UndoRecord.CustomRecordName.md)**|Property|Returns a **String** that specifies the entry that appears on the undo stack when all custom undo actions have completed. Read-only.|
|**[IsRecordingCustomRecord](../../../api/Word.UndoRecord.IsRecordingCustomRecord.md)**|Property|Returns a **Boolean** that specifies whether a custom undo action is being recorded. Read-only.|

## Best Practices


- Close each custom undo record with a call to [EndCustomRecord](../../../api/Word.UndoRecord.EndCustomRecord.md). Word will attempt to determine where to end the record, but it may not be at the desired point of code execution.
    
- Call the [Document](../../../api/Word.Document.md) object's [Undo](../../../api/Word.Document.Undo.md) method in the right order when calling from within a custom undo record. Calling the [Undo](../../../api/Word.Document.Undo.md) method in the wrong order within a custom undo record can have undesired effects on your code execution. The following code example shows the [Undo](../../../api/Word.Document.Undo.md) method being called from inside a custom undo record.
    
```vb
Sub UndoInUndoRecord() 
   
  Set objUndo = Application.UndoRecord 
  
  objUndo.StartCustomRecord ("New Paragraph") 
  
  ActiveDocument.Content.InsertAfter Text:=" The end." 
  ActiveDocument.Undo 
  ActiveDocument.Content.InsertAfter Text:=" The end, again." 
  ActiveDocument.Content.InsertAfter Text:=" The last end." 
 
  objUndo.EndCustomRecord 
 
End Sub
```

When this code is run, each action in the custom undo record prior to the **Undo** method call is undone. Each action in the custom undo record after the **Undo** method call is placed the undo stack.

- Avoid switching documents inside a custom undo record. The following code example switches from one document to another inside a custom undo record.

```vb
Dim objUndo As UndoRecord 

Sub SwitchDocsInsideUndo() 
  Set objUndo = Application.UndoRecord objUndo.StartCustomRecord ("New Paragraph") 
  
  'Insert some text into the first document Documents(1).Content.InsertAfter "A new paragraph in doc1." 
  
  'Switch documents to the second document 
  'The custom undo record will terminate here Documents(2).Content.InsertAfter "A new paragraph in doc2." 
  
  objUndo.EndCustomRecord 
End Sub
```

When this code is run, Word terminates the custom undo record when the code begins to write to the second document. When the procedure is finished, only the first document will have a custom undo record with the name "New Paragraph."
    
- Avoid using Breaks when running code that contains custom undo records in Debug mode. When code that contains undo records is run in Debug mode in the Visual Basic editor and the code encounters a breakpoint, Word automatically ends all custom records that are currently in progress.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]