---
title: Fields.AddHorizontalInVertical method (Publisher)
keywords: vbapb10.chm6029319
f1_keywords:
- vbapb10.chm6029319
ms.prod: publisher
api_name:
- Publisher.Fields.AddHorizontalInVertical
ms.assetid: 4b451a24-0d79-70d4-4910-2725f1ed0297
ms.date: 06/07/2019
localization_priority: Normal
---


# Fields.AddHorizontalInVertical method (Publisher)

Inserts horizontal text into a stream of vertical text and returns the new horizontal text as a **[Field](Publisher.Field.md)** object.


## Syntax

_expression_.**AddHorizontalInVertical** (_Range_, _Text_)

_expression_ A variable that represents a **[Fields](Publisher.Fields.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Range_|Required| **TextRange**|The text range at which to insert the horizontal text.|
|_Text_|Required| **String**|The text to be horizontally inserted.|

## Return value

Field


## Example

This example horizontally inserts the text Horizontal Test after the existing vertical text in shape one on page one of the active publication.

```vb
Dim rngTemp As TextRange 
Dim fldTemp As Field 
 
With ActiveDocument.Pages(1).Shapes(1) 
 Set rngTemp = .TextFrame.TextRange.InsertAfter("") 
 
 Set fldTemp = .TextFrame.TextRange.Fields _ 
 .AddHorizontalInVertical(Range:=rngTemp, Text:="Horizontal Test") 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]