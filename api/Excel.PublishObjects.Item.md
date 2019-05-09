---
title: PublishObjects.Item property (Excel)
keywords: vbaxl10.chm650075
f1_keywords:
- vbaxl10.chm650075
ms.prod: excel
api_name:
- Excel.PublishObjects.Item
ms.assetid: 5327f5b3-8dd0-cb10-49b5-9824d0376667
ms.date: 05/09/2019
localization_priority: Normal
---


# PublishObjects.Item property (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[PublishObjects](Excel.PublishObjects.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number of the object.|

## Example

This example obtains the identifier from a `<DIV>` tag and finds the line in a webpage (q198.htm) that you saved from a workbook. The example then creates a copy of the webpage (newq1.htm) and inserts a comment line before the `<DIV>` tag in the copy of the file.

```vb
strTargetDivID = ActiveWorkbook.PublishObjects.Item(1).DivID 
Open "\\server1\reports\q198.htm" For Input As #1 
Open "\\server1\reports\newq1.htm" For Output As #2 
While Not EOF(1) 
 Line Input #1, strFileLine 
 If InStr(strFileLine, strTargetDivID) > 0 And _ 
 InStr(strFileLine, "<div") > 0 Then 
 Print #2, "<!--Saved item-->" 
 End If 
 Print #2, strFileLine 
Wend 
Close #2 
Close #1
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]