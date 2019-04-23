---
title: Explorer.ActiveInlineResponseWordEditor property (Outlook)
keywords: vbaol11.chm3597
f1_keywords:
- vbaol11.chm3597
ms.assetid: b9058694-ab8f-4962-ab7d-afac1704dd29
ms.date: 06/08/2017
ms.prod: outlook
localization_priority: Normal
---


# Explorer.ActiveInlineResponseWordEditor property (Outlook)
Returns the Word [Document](./Word.Document.md) object of the active inline response that is displayed in the explorer Reading Pane. Read-only.

## Syntax

_expression_. `ActiveInlineResponseWordEditor`

_expression_ A variable that represents an '[Explorer](Outlook.Explorer.md)' object.


## Remarks

This property returns  **Null** (**Nothing** in Visual Basic) if no inline response is visible in the Reading Pane. The returned Word **Document** object provides access to most of the Word object model except for the following members:


- [InlineShapes.AddChart2](./Word.inlineshapes.addchart2.md)
    
- [Range.ConvertToTable](./Word.Range.ConvertToTable.md)
    
- [Range.ImportFragment](./Word.Range.ImportFragment.md)
    
- [Range.InsertXML](./Word.Range.InsertXML.md)
    
- [Shapes.AddChart2](./Word.shapes.addchart2.md)
    
- [Selection.InsertXML](./Word.Selection.InsertXML.md)
    
- [Tables.Add](./Word.Tables.Add.md)
    

## See also


[Explorer Object](Outlook.Explorer.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]