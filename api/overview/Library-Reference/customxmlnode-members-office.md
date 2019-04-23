---
title: CustomXMLNode members (Office)
ms.prod: office
ms.assetid: fbf957c8-40b8-2f75-fcc8-db0ed6e18438
ms.date: 01/30/2019
localization_priority: Normal
---


# CustomXMLNode members (Office)

Represents an XML node in a tree in a document. The **CustomXMLNode** object is a member of the **CustomXMLNodes** collection.


## Methods

|Name|Description|
|:-----|:-----|
|[AppendChildNode](../../Office.CustomXMLNode.AppendChildNode.md)|Appends a single node as the last child under the context element node in the tree. |
|[AppendChildSubtree](../../Office.CustomXMLNode.AppendChildSubtree.md)|Adds a subtree as the last child under the context element node in the tree.|
|[Delete](../../Office.CustomXMLNode.Delete.md)|Deletes the current node from the tree (including all of its children, if any exist).|
|[HasChildNodes](../../Office.CustomXMLNode.HasChildNodes.md)|Gets **True** if the current element node has child element nodes.|
|[InsertNodeBefore](../../Office.CustomXMLNode.InsertNodeBefore.md)|Inserts a new node just before the context node in the tree.|
|[InsertSubtreeBefore](../../Office.CustomXMLNode.InsertSubtreeBefore.md)|Inserts the specified subtree into the location just before the context node. |
|[RemoveChild](../../Office.CustomXMLNode.RemoveChild.md)|Removes the specified child node from the tree.|
|[ReplaceChildNode](../../Office.CustomXMLNode.ReplaceChildNode.md)|Removes the specified child node (and its subtree) from the main tree, and replaces it with a different node in the same location.|
|[ReplaceChildSubtree](../../Office.CustomXMLNode.ReplaceChildSubtree.md)|Removes the specified node (and its subtree) from the main tree, and replaces it with a different subtree in the same location.|
|[SelectNodes](../../Office.CustomXMLNode.SelectNodes.md)|Selects a collection of nodes matching an XPath expression. This method differs from the **CustomXMLPart**. **SelectNodes** method in that the XPath expression will be evaluated starting with the 'expression' node as the context node.|
|[SelectSingleNode](../../Office.CustomXMLNode.SelectSingleNode.md)|Selects a single node from a collection matching an XPath expression. This method differs from the **CustomXMLPart**. **SelectSingleNode** method in that the XPath expression will be evaluated starting with the 'expression' node as the context node.|


## Properties

|Name|Description|
|:-----|:-----|
|[Application](../../Office.CustomXMLNode.Application.md)|Gets an **Application** object that represents the container application for a **CustomXMLNode**. Read-only.|
|[Attributes](../../Office.CustomXMLNode.Attributes.md)|Gets a **CustomXMLNodes** collection representing the attributes of the current element in the current node. Read-only.|
|[BaseName](../../Office.CustomXMLNode.BaseName.md)|Gets the base name of the node without the namespace prefix, if one exists, in the Document Object Model (DOM). Read-only.|
|[ChildNodes](../../Office.CustomXMLNode.ChildNodes.md)|Gets a **CustomXMLNodes** collection containing all of the child elements of the current node. Read-only.|
|[Creator](../../Office.CustomXMLNode.Creator.md)|Gets a 32-bit integer that indicates the application in which the **CustomXMLNode** object was created. Read-only.|
|[FirstChild](../../Office.CustomXMLNode.FirstChild.md)|Gets a **CustomXMLNode** object corresponding to the first child element of the current node. If the node has no child elements (or if it isn't of type **msoCustomXMLNodeElement**), returns **Nothing**. Read-only.|
|[LastChild](../../Office.CustomXMLNode.LastChild.md)|Gets a **CustomXMLNode** object corresponding to the last child element of the current node. If the node has no child elements (or if it is not of type **msoCustomXMLNodeElement**), the property returns **Nothing**. Read-only.|
|[NamespaceURI](../../Office.CustomXMLNode.NamespaceURI.md)|Gets the unique address identifier for the namespace of the **CustomXMLNode** object. Read-only.|
|[NextSibling](../../Office.CustomXMLNode.NextSibling.md)|Gets the next sibling node (element, comment, or processing instruction) of the current node. If the node is the last sibling at its level, the property returns **Nothing**. Read-only.|
|[NodeType](../../Office.CustomXMLNode.NodeType.md)|Gets the type of the current node. Read-only.|
|[NodeValue](../../Office.CustomXMLNode.NodeValue.md)|Gets or sets the value of the current node. Read/write.|
|[OwnerDocument](../../Office.CustomXMLNode.OwnerDocument.md)|Gets the object representing the Microsoft Excel workbook, Microsoft PowerPoint presentation, or the Microsoft Word document associated with this node. Read-only.|
|[OwnerPart](../../Office.CustomXMLNode.OwnerPart.md)|Gets the object representing the part associated with this node. Read-only.|
|[Parent](../../Office.CustomXMLNode.Parent.md)|Gets the **Parent** object for the **CustomXMLNode** object. Read-only.|
|[ParentNode](../../Office.CustomXMLNode.ParentNode.md)|Gets the parent element node of the current node. If the current node is at the root level, the property returns **Nothing**. Read-only.|
|[PreviousSibling](../../Office.CustomXMLNode.PreviousSibling.md)|Gets the previous sibling node (element, comment, or processing instruction) of the current node. If the current node is the first sibling at its level, the property returns **Nothing**. Read-only.|
|[Text](../../Office.CustomXMLNode.Text.md)|Gets or sets the text for the current node. Read/write.|
|[XML](../../Office.CustomXMLNode.XML.md)|Gets the XML representation of the current node and its children, if any exist. Read-only.|
|[XPath](../../Office.CustomXMLNode.XPath.md)|Gets a **String** with the canonicalized XPath for the current node. If the node is no longer in the Document Object Model (DOM), the property returns an error message. Read-only.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]