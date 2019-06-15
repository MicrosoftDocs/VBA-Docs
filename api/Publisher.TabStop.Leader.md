---
title: TabStop.Leader property (Publisher)
keywords: vbapb10.chm5636101
f1_keywords:
- vbapb10.chm5636101
ms.prod: publisher
api_name:
- Publisher.TabStop.Leader
ms.assetid: a788bdc8-8ab3-fcd3-931a-a5b83db93991
ms.date: 06/15/2019
localization_priority: Normal
---


# TabStop.Leader property (Publisher)

Sets or returns a **[PbTabLeaderType](Publisher.PbTabLeaderType.md)** constant that represents the leader character for a tab stop. Read/write.


## Syntax

_expression_.**Leader**

_expression_ A variable that represents a **[TabStop](Publisher.TabStop.md)** object.


## Return value

PbTabLeaderType


## Remarks

The **Leader** property value can be one of the **PbTabLeaderType** constants declared in the Microsoft Publisher type library.


## Example

This example changes the leader tab character of the selected paragraphs to dashes. This example assumes that the selected paragraph contains at least one tab stop.

```vb
Sub SetLeaderTab() 
 Selection.TextRange.ParagraphFormat _ 
 .Tabs(1).Leader = pbTabLeaderDashes 
End Sub
```

<br/>

This example changes the leader tab character of the first paragraph in the specified text range to an underline. This example assumes that the specified paragraph contains at least one tab stop.

```vb
Sub SetNewTabLeader() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange.Paragraphs(1) _ 
 .ParagraphFormat.Tabs(1).Leader = pbTabLeaderLine 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]