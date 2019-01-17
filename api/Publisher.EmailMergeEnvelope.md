---
title: EmailMergeEnvelope Object (Publisher)
keywords: vbapb10.chm9109503
f1_keywords:
- vbapb10.chm9109503
ms.prod: publisher
api_name:
- Publisher.EmailMergeEnvelope
ms.assetid: 555dd80e-bac2-96dd-4256-ad1b8006da0f
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailMergeEnvelope Object (Publisher)

Represents the email container (envelope) that holds the Microsoft Publisher document that is merged into an email merge.
 


## Remarks

The properties of the  **EmailMergeEnvelope** object correspond to the combination of both required and optional settings in the **Merge to Email** dialog box in the Publisher user interface (on the **File** menu, point to **Send Email**, click  **Send Email Merge**, and then click  **Options**). 
 

 
Before you can use the  **Execute** method of the **[MailMerge](Publisher.MailMerge.md)** object to send a merged email, you must specify a value for the **To** property of the **EmailMergeEnvelope** object, or Publisher will return an error.
 

 

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to assign some of the properties of an  **EmailMergeEnvelope** object that represents an email merge and then send the resulting email message, an invitation. The macro connects to a data source, assigns values to the **To** and **Subject** properties of the **EmailMergeEnvelope** object, and adds a text box containing merge fields and some additional text to the email message. Then it uses the **Execute** method of the **MailMerge** object to execute the merge and send the email.
 

 
The data source referenced in this example is a simple tab-delimited text file that contains three columns with the headings "First," "Last," and "Email Address" respectively.
 

 
Before running the code, create the text file, add one or more data rows, name the file DataSource.txt, and save it to disk. Then add the file's path to the code by replacing the  _PathToFile_ variable with your path.
 

 
If you run the code in this example more than once, you will encounter errors because Publisher connects to the data source each time you run the code, resulting in a publication connected to multiple data sources. When multiple data-source connections exist, Publisher inserts an extra column in the master (combined) mail-merge data source to specify the specific data source for each record. As a result, Publisher effectively changes the index number of all the data-source columns, making the indexes used in this code (for example,  _MailMergeField1_ ) incorrect.
 

 



```vb
Public Sub EmailMergeEnvelope_Example() 
 
 Dim pubShape As Publisher.Shape 
 Dim pubMailMerge As Publisher.MailMerge 
 
 'Connect to the data source. 
 Set pubMailMerge = ThisDocument.MailMerge 
 pubMailMerge.OpenDataSource "PathToFile \DataSource.txt" 
 
 'Assign "Email Address" to the To field of the email message. 
 pubMailMerge.EmailMergeEnvelope.To = pubMailMerge.DataSource.DataFields.Item(3) 
 
 'Add text to the Subject field of the email message. 
 pubMailMerge.EmailMergeEnvelope.Subject = "Invitation" 
 
 'Insert two merge fields and some additional text in a text box in the body of the message. 
 Set pubShape = ThisDocument.Pages(1).Shapes.AddTextbox(pbTextOrientationHorizontal, 100, 100, 200, 100) 
 pubShape.TextFrame.TextRange.Text = "Dear " 
 pubShape.TextFrame.TextRange.InsertMailMergeField 1 
 pubShape.TextFrame.TextRange.InsertAfter " " 
 pubShape.TextFrame.TextRange.InsertMailMergeField 2 
 pubShape.TextFrame.TextRange.InsertAfter ": " 
 pubShape.TextFrame.TextRange.InsertAfter "You are invited!" 
 
 'Perform the merge. 
 pubMailMerge.Execute True, pbSendEmail 
 
 'Display a reminder 
 MsgBox "If your email client is not already open, remember to open it and send the email messages that are in the outbox." 
 
End Sub
```


## Properties



|Name|
|:-----|
|[Application](Publisher.EmailMergeEnvelope.Application.md)|
|[Attachments](Publisher.EmailMergeEnvelope.Attachments.md)|
|[Bcc](Publisher.EmailMergeEnvelope.Bcc.md)|
|[Cc](Publisher.EmailMergeEnvelope.Cc.md)|
|[Parent](Publisher.EmailMergeEnvelope.Parent.md)|
|[Priority](Publisher.EmailMergeEnvelope.Priority.md)|
|[Subject](Publisher.EmailMergeEnvelope.Subject.md)|
|[To](Publisher.EmailMergeEnvelope.To.md)|

