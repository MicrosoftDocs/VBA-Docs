---
title: PublishObject.Publish method (Excel)
keywords: vbaxl10.chm652074
f1_keywords:
- vbaxl10.chm652074
ms.prod: excel
api_name:
- Excel.PublishObject.Publish
ms.assetid: 3bb70102-c440-8e49-1734-d72945324d5c
ms.date: 05/09/2019
localization_priority: Normal
---


# PublishObject.Publish method (Excel)

Saves an item or a collection of items in a document to a webpage.


## Syntax

_expression_.**Publish** (_Create_)

_expression_ A variable that represents a **[PublishObject](Excel.PublishObject.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Create_|Optional| **Variant**|If the HTML file exists, setting this argument to **True** replaces the file, and setting this argument to **False** inserts the item or items at the end of the file. If the file does not exist, the file is created regardless of the value of the _Create_ argument.|

## Remarks

The **[FileName](Excel.PublishObject.Filename.md)** property returns or sets the location and name of the HTML file.


## Example

This example saves the range D5:D9 on the First Quarter worksheet in the active workbook to a webpage named Stockreport.htm. The spreadsheet component is used to make the webpage interactive.

```vb
With ActiveWorkbook.PublishObjects.Add(xlSourceRange, _ 
 "\\Server1\sharedfolder\Stockreport.htm", "First Quarter", _ 
 "$D$5:$D$9", xlHtmlStatic, "Book2_25082", "") 
 .Publish (True) 
 .AutoRepublish = True 
End With
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]