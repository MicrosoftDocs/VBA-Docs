---
title: MailMerge.CreateShortcut method (Publisher)
keywords: vbapb10.chm6225942
f1_keywords:
- vbapb10.chm6225942
ms.prod: publisher
api_name:
- Publisher.MailMerge.CreateShortcut
ms.assetid: 96878925-41ce-4873-931e-d5c05307a94a
ms.date: 06/08/2019
localization_priority: Normal
---


# MailMerge.CreateShortcut method (Publisher)

Creates a shortcut to the file that contains the list of recipients or products for a mail merge publication.


## Syntax

_expression_.**CreateShortcut** (_FileName_)

_expression_ A variable that represents a **[MailMerge](Publisher.MailMerge.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_FileName_|Required| **String**|The name of the mailing list or product list file for which the shortcut should be created.|

## Remarks

The **CreateShortcut** method corresponds to the **Save a shortcut to recipient list** command in the **Mail Merge** and **Email Merge** task panes, and the **Save a shortcut to product list** command in the **Catalog Merge** task pane in the Microsoft Publisher user interface.

Mailing list recipient files have the .ols extension (for Microsoft Office List Shortcut).


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **CreateShortcut** method to create a shortcut to a mail-merge recipient list. Before running this macro, ensure that the active document is connected to a data source. If the active document is not already connected to a data source, you can use the **[OpenDataSource](Publisher.MailMerge.OpenDataSource.md)** method to make the connection.

Also, before running the code, replace _username_ in the folder path to the saved file with the name of a valid user on your computer, or replace the folder path and file name with a path and file name of your choice.

Note that the folder path used in this example is typical of folder paths in Windows. You must have permission to save files in the folder that you designate.

```vb
Public Sub CreateShortcut_Example() 
 
 Dim pubMailMerge As Publisher.MailMerge 
 Set pubMailMerge = ThisDocument.MailMerge 
 
 pubMailMerge.CreateShortcut ("C:\Users\username\Documents\My Data Sources\MyRecipientList") 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]