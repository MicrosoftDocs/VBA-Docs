---
title: Chart.Export method (Word)
keywords: vbawd10.chm79364170
f1_keywords:
- vbawd10.chm79364170
ms.prod: word
api_name:
- Word.Chart.Export
ms.assetid: 49660450-ae9f-c59e-8974-b04327a72dc0
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.Export method (Word)

Exports the chart in a graphic format.


## Syntax

_expression_.**Export** (_FileName_, _FilterName_, _Interactive_)

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The name of the exported file.|
| _FilterName_|Optional| **Variant**|The language-independent name of the graphic filter as it appears in the registry.|
| _Interactive_|Optional| **Variant**| **True** to display the dialog box that contains the filter-specific options. **False** to indicate that Word should use the default values for the filter. The default is **False**.|

## Return value

A  **Boolean** value that indicates whether the export was successful.


## Example

The following example exports the first chart in the active document as a GIF file.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Export _ 
 FileName:="current_sales.gif", FilterName:="GIF" 
 End If 
End With
```


## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]