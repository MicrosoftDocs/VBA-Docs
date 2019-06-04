---
title: Application.MailMergeAfterRecordMerge event (Publisher)
keywords: vbapb10.chm268435472
f1_keywords:
- vbapb10.chm268435472
ms.prod: publisher
api_name:
- Publisher.Application.MailMergeAfterRecordMerge
ms.assetid: 550c3310-01ba-718f-4c1d-cbf3ce077d27
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.MailMergeAfterRecordMerge event (Publisher)

Occurs after each record in the data source successfully merges in a mail merge.


## Syntax

_expression_.**MailMergeAfterRecordMerge** (_Doc_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Doc_|Required| **Document**|The mail merge main document.|

## Remarks

If you maintain a customer management database, you can use the **MailMergeAfterRecordMerge** event to update the database for each merged record.

To access the **Application** object events, declare an **Application** object variable in the General Declarations section of a code module, and then set the variable equal to the **Application** object for which you want to access events. 

For information about using events with the Microsoft Publisher **Application** object, see [Using events with the Application object](../publisher/Concepts/using-events-with-the-application-object-publisher.md).


## Example

This example displays a message with the value of the first and second fields in the record that has just finished merging.

```vb
Private Sub MailMergeApp_MailMergeAfterRecordMerge(ByVal Doc As Document) 
 
 With ActiveDocument.MailMerge.DataSource 
 MsgBox .DataFields.Item(3).Value & " " & _ 
 .DataFields.Item(2).Value & " is finished merging." 
 End With 
 
End Sub
```

<br/>

For this event to occur, you must place the following line of code in the General Declarations section of your module and run the following initialization routine.

```vb
Private WithEvents MailMergeApp As Application 
 
Sub InitializeMailMergeApp() 
 Set MailMergeApp = Publisher.Application 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]