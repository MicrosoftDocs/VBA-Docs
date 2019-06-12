---
title: MailMerge.OpenDataSource method (Publisher)
keywords: vbapb10.chm6225937
f1_keywords:
- vbapb10.chm6225937
ms.prod: publisher
api_name:
- Publisher.MailMerge.OpenDataSource
ms.assetid: 4473e566-687f-595e-9fd6-a5483021cb48
ms.date: 06/08/2019
localization_priority: Normal
---


# MailMerge.OpenDataSource method (Publisher)

Attaches a data source to the specified publication, which becomes a main publication if it is not one already.


## Syntax

_expression_.**OpenDataSource** (_bstrDataSource_, _bstrConnect_, _bstrTable_, _fOpenExclusive_, _fNeverPrompt_)

_expression_ A variable that represents a **[MailMerge](publisher.mailmerge.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_bstrDataSource_|Optional| **String**|The data source path and file name. You can specify a Microsoft Query (.qry) file instead of specifying a data source, a connection string, and a table name string; values in a Microsoft Query file override values for _bstrConnect_ and _bstrTable_.|
|_bstrConnect_|Optional| **String**|A connection string.|
|_bstrTable_|Optional| **String**|The name of the table in the data source.|
|_fOpenExclusive_|Optional| **Long**| **True** to deny others access to the database. **False** allows others read/write permission to the database. The default value is **False**.|
|_fNeverPrompt_|Optional| **Long**| **True** never prompts when opening the data source. **False** displays the **Data Link Properties** dialog box. The default value is **False**.|

## Remarks

If you are using a data source for mail merge, you must add a catalog merge area to the publication page before you attach to the data source.


## Example

This example attaches a table from a database and denies everyone else write access to the database while it is opened. 

For this example to run properly, you must replace `PathToFile` with a valid file path and `TableName` with a valid data source table name.

```vb
Sub AttachDataSource() 
 
    ActiveDocument.MailMerge.OpenDataSource _ 
        bstrDataSource:="PathToFile",  _ 
        bstrTable:="TableName", _ 
        fNeverPrompt:=True, fOpenExclusive:=True 
 
End Sub
```

> [!NOTE] 
> For `TableName`, if an Excel spreadsheet is being opened, `TableName` must be followed by `$`. That is, `bstrTable:="Sheet1"` will not work; `bstrTable:="Sheet1$"` will work. Following is an example that further clarifies this.

<br/>

In this example, the data is stored in MySpreadSheet.xlsx, Sheet1, in the same directory as the Publisher file.

```vb
Dim strDataFile as String
strDataFile = Application.ActiveDocument.Path & "MySpreadSheet.xlsx"

ActiveDocument.MailMerge.OpenDataSource _ 
    bstrDataSource:=strDataFile,  _ 
    bstrTable:="Sheet1$", _ 
    fNeverPrompt:=True, fOpenExclusive:=True 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]