---
title: TextRange.InsertSymbol method (Publisher)
keywords: vbapb10.chm5308452
f1_keywords:
- vbapb10.chm5308452
ms.prod: publisher
api_name:
- Publisher.TextRange.InsertSymbol
ms.assetid: 607d12da-5a2d-4e0e-b45e-92275ce97bab
ms.date: 06/15/2019
localization_priority: Normal
---


# TextRange.InsertSymbol method (Publisher)

Returns a **TextRange** object that represents a symbol inserted in place of the specified range or selection.


## Syntax

_expression_.**InsertSymbol** (_FontName_, _CharIndex_)

_expression_ A variable that represents a **[TextRange](Publisher.TextRange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_FontName_|Required| **String**|The name of the font that contains the symbol.|
|_CharIndex_|Required| **Long**|The Unicode character for the specified symbol.|

## Return value

TextRange


## Remarks

If you do not want to replace the range or selection, use the **[Collapse](Publisher.TextRange.Collapse.md)** method before you use this method.


## Example

This example inserts a double-headed arrow at the cursor.

```vb
Sub Insert Arrow() 
    ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
            .Paragraphs(Start:=1, Length:=1).Select
    With .TextFrame.TextRange 
            .InsertPageNumber 
            .Collapse Direction:= pbCollapseStart
            .InsertSymbol FontName:="Symbol", CharIndex:=171
        End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]