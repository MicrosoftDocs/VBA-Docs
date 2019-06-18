---
title: WebCommandButton.DataFileName property (Publisher)
keywords: vbapb10.chm3932165
f1_keywords:
- vbapb10.chm3932165
ms.prod: publisher
api_name:
- Publisher.WebCommandButton.DataFileName
ms.assetid: 5fd2bac7-7067-4833-4b34-26897c39ea58
ms.date: 06/18/2019
localization_priority: Normal
---


# WebCommandButton.DataFileName property (Publisher)

Returns or sets a **String** that represents the name of the file in which to save data from a web form. Read/write.


## Syntax

_expression_.**DataFileName**

_expression_ A variable that represents a **[WebCommandButton](Publisher.WebCommandButton.md)** object.


## Return value

String


## Example

This example sets Microsoft Publisher to process web form data by saving it to a comma-delimited text file on the same web server as the form is stored.

```vb
Sub WebDataFile() 
 With ThisDocument.Pages(1).Shapes(1).WebCommandButton 
 .DataRetrievalMethod = pbSubmitDataRetrievalSaveOnServer 
 .DataFileFormat = pbSubmitDataFormatCSV 
 .DataFileName = "WebFormData.txt" 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]