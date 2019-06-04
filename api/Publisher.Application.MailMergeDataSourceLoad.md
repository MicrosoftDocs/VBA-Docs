---
title: Application.MailMergeDataSourceLoad event (Publisher)
keywords: vbapb10.chm268435475
f1_keywords:
- vbapb10.chm268435475
ms.prod: publisher
api_name:
- Publisher.Application.MailMergeDataSourceLoad
ms.assetid: afca3a05-d6a6-15f1-8cbf-593777066757
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.MailMergeDataSourceLoad event (Publisher)

Occurs when the data source is loaded for a mail merge.


## Syntax

_expression_.**MailMergeDataSourceLoad** (_Doc_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Doc_|Required| **Document**|The mail merge main document.|

## Remarks

To access the **Application** object events, declare an **Application** object variable in the General Declarations section of a code module, and then set the variable equal to the **Application** object for which you want to access events. 

For information about using events with the Microsoft Publisher **Application** object, see [Using events with the Application object](../publisher/Concepts/using-events-with-the-application-object-publisher.md).


## Example

This example displays a message with the data source file name when the data source starts loading.

```vb
Private Sub MailMergeApp_MailMergeDataSourceLoad(ByVal Doc As Document) 
 Dim strDSName As String 
 Dim intDSLength As Integer 
 Dim intDSStart As Integer 
 
 'Pull out of the Name property (which includes path and file name) 
 'only the file name using Visual Basic commands Len, InStrRev, and Right 
 intDSLength = Len(ActiveDocument.MailMerge.DataSource.Name) 
 intDSStart = InStrRev(ActiveDocument.MailMerge.DataSource.Name, "\") 
 intDSStart = intDSLength - intDSStart 
 strDSName = Right(ActiveDocument.MailMerge.DataSource.Name, intDSStart) 
 
 'Deliver a message to user when data source is loading 
 MsgBox "Your data source, " & strDSName & ", is now loading." 
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