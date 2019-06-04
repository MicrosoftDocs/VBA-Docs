---
title: Application.MailMergeAfterMerge event (Publisher)
keywords: vbapb10.chm268435465
f1_keywords:
- vbapb10.chm268435465
ms.prod: publisher
api_name:
- Publisher.Application.MailMergeAfterMerge
ms.assetid: dd01d8f5-f95e-e833-bb8b-708ced54240c
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.MailMergeAfterMerge event (Publisher)

Occurs after all records in a mail merge have merged successfully.


## Syntax

_expression_.**MailMergeAfterMerge** (_Doc_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Doc_|Required| **Document**|The mail merge main document.|

## Remarks

To access the **Application** object events, declare an **Application** object variable in the General Declarations section of a code module, and then set the variable equal to the **Application** object for which you want to access events. 

For information about using events with the Microsoft Publisher **Application** object, see [Using events with the Application object](../publisher/Concepts/using-events-with-the-application-object-publisher.md).


## Example

This example displays a message stating that all records in the specified document are finished merging.

```vb
Private Sub MailMergeApp_MailMergeAfterMerge(ByVal Doc As Document) 
 
 MsgBox "Your mail merge on " & _ 
 ActiveDocument.Name & " is now finished." 
 
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