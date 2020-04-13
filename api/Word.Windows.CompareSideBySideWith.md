---
title: Windows.CompareSideBySideWith method (Word)
keywords: vbawd10.chm157351948
f1_keywords:
- vbawd10.chm157351948
ms.prod: word
api_name:
- Word.Windows.CompareSideBySideWith
ms.assetid: 522c75b2-460a-460f-93ef-71cc84973d2f
ms.date: 06/08/2017
localization_priority: Normal
---


# Windows.CompareSideBySideWith method (Word)

Opens two windows in side by side mode. Returns a **Boolean**.


## Syntax

_expression_. `CompareSideBySideWith`( `_Document_` )

_expression_ Required. A variable that represents a '[Windows](Word.windows.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Document_|Required| **Variant**| The document to view in side by side windows.|

## Return value

Boolean


## Remarks

You cannot use the **CompareSideBySideWith** method with the **Application** object or the **ActiveDocument** property.


## Example

The following example places two new documents in adjacent windows.


```vb
Dim objDoc1 As Word.Document 
Dim objDoc2 As Word.Document 
 
Set objDoc1 = Documents.Add 
Set objDoc2 = Documents.Add 
 
objDoc2.Activate 
objDoc2.Windows.CompareSideBySideWith objDoc1 
Windows.ResetPositionsSideBySide
```


## See also


[Windows Collection Object](Word.windows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]