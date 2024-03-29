---
title: Hyperlink.CreateNewDocument method (Excel)
keywords: vbaxl10.chm536086
f1_keywords:
- vbaxl10.chm536086
api_name:
- Excel.Hyperlink.CreateNewDocument
ms.assetid: 902914b7-08ea-0839-13e1-8fc7e7192675
ms.date: 04/26/2019
ms.localizationpriority: medium
---


# Hyperlink.CreateNewDocument method (Excel)

Creates a new document linked to the specified hyperlink.


## Syntax

_expression_.**CreateNewDocument** (_FileName_, _EditNow_, _Overwrite_)

_expression_ A variable that represents a **[Hyperlink](Excel.Hyperlink.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The file name of the specified document.|
| _EditNow_|Required| **Boolean**| **True** to have the specified document open immediately in its associated editing environment. The default value is **True**.|
| _Overwrite_|Required| **Boolean**| **True** to overwrite any existing file of the same name in the same folder. **False** if any existing file of the same name is preserved and the _FileName_ argument specifies a new file name. The default value is **False**.|

## Example

This example creates a new document based on the new hyperlink in the first worksheet and then loads the document into Microsoft Excel for editing. The document is called Report.xls, and it overwrites any file of the same name in the \\Server1\Annual folder.

```vb
With Worksheets(1) 
 Set objHyper = _ 
 .Hyperlinks.Add(Anchor:=.Range("A10"), _ 
 Address:="\\Server1\Annual\Report.xls") 
 objHyper.CreateNewDocument _ 
 FileName:="\\Server1\Annual\Report.xls", _ 
 EditNow:=True, Overwrite:=True 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]