---
title: SmartArtNode members (Office)
description: A single semantic node within the data model of a SmartArt graphic.
ms.prod: office
ms.assetid: 8472d586-87ed-2dd7-054b-e821f1738e3c
ms.date: 01/30/2019
localization_priority: Normal
---


# SmartArtNode members (Office)

A single semantic node within the data model of a SmartArt graphic.


## Methods

|Name|Description|
|:-----|:-----|
|[AddNode](../../Office.SmartArtNode.AddNode.md)|Adds a new **SmartArtNode** to the data model in the way specified by the **SmartArtNodePosition** value, and of type **SmartArtNodeType**.|
|[Delete](../../Office.SmartArtNode.Delete.md)|Removes the current SmartArt node. |
|[Demote](../../Office.SmartArtNode.Demote.md)|Demotes the current node a single level within the data model.|
|[Larger](../../Office.SmartArtNode.Larger.md)|Increases the size of the SmartArt node. Mimics the behavior of the **Larger** button on the Microsoft Office Fluent Ribbon **Format** tab for SmartArt.|
|[Promote](../../Office.SmartArtNode.Promote.md)|Promotes the current node (and all its children) a single level within the data model.|
|[ReorderDown](../../Office.SmartArtNode.ReorderDown.md)|Swaps a node with the next node in the bulleted list. This method reorder's the nodes entire family.|
|[ReorderUp](../../Office.SmartArtNode.ReorderUp.md)|Swaps a node with the previous node in the bulleted list. This method reorder's the nodes entire family.|
|[Smaller](../../Office.SmartArtNode.Smaller.md)|Decreases the size of the SmartArt. Mimics the behavior of the **Smaller** button on the Microsoft Office Fluent Ribbon UI **Format** tab for SmartArt.|

## Properties

|Name|Description|
|:-----|:-----|
|[Application](../../Office.SmartArtNode.Application.md)|Gets an **Application** object that represents the container application for the **SmartArtNode** object. Read-only.|
|[Creator](../../Office.SmartArtNode.Creator.md)|Gets a 32-bit integer that indicates the application in which the **SmartArtNode** object was created. Read-only.|
|[Hidden](../../Office.SmartArtNode.Hidden.md)|Returns **True** if this node is a hidden node in the data model. Read-only.|
|[Level](../../Office.SmartArtNode.Level.md)|Retrieves the node's level in the hierarchy. Read-only.|
|[Nodes](../../Office.SmartArtNode.Nodes.md)|Retrieves the children nodes associated with this **SmartArtNode**. Read-only.|
|[OrgChartLayout](../../Office.SmartArtNode.OrgChartLayout.md)|Retrieves or sets the **MsoOrgChartLayoutType** associated with this node if there is one. Read/write.|
|[Parent](../../Office.SmartArtNode.Parent.md)|Returns the calling object. Read-only.|
|[ParentNode](../../Office.SmartArtNode.ParentNode.md)|Retrieves the parent **SmartArtNode** of this **SmartArtNode**. Read-only.|
|[Shapes](../../Office.SmartArtNode.Shapes.md)|Returns the shape range associated with this **SmartArtNode** object. Read-only.|
|[TextFrame2](../../Office.SmartArtNode.TextFrame2.md)|Returns the text associated with the **SmartArtNode** object. Read-only.|
|[Type](../../Office.SmartArtNode.Type.md)|Retrieves the type of **SmartArtNode** object. Read-only.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]