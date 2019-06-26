---
title: GraphicItem.Index property (Visio)
keywords: vis_sdr.chm16913695
f1_keywords:
- vis_sdr.chm16913695
ms.prod: visio
api_name:
- Visio.GraphicItem.Index
ms.assetid: 44dde969-4330-8ad0-5ed2-a80e4c755143
ms.date: 06/08/2017
localization_priority: Normal
---


# GraphicItem.Index property (Visio)

Gets or sets the ordinal position of a **GraphicItem** object in the **GraphicItems** collection of a data graphic master—a **Master** object of type **visTypeDataGraphic**. Read/write.


> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

_expression_.**Index**

_expression_ A variable that represents a **[GraphicItem](Visio.GraphicItem.md)** object.


## Return value

Long


## Remarks

The index of a graphic item is originally determined by the order in which the item was added to the collection. The **GraphicItems** collection is 1-based.

The index order of graphic items affects the stacking order for multiple graphic item callouts assigned to the same location. In addition, it determines which graphic item takes precedence in control over a cell in the Microsoft Visio ShapeSheet spreadsheet when conflicting conditions set by multiple graphic items are all true .


> [!NOTE] 
> Before you can set any property of a graphic item, you must use the **[Master.Open](Visio.Master.Open.md)** method to open a copy of the data graphic master that contains the graphic item for editing. When you are finished setting properties, use the **Master.Close** method to commit changes.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]