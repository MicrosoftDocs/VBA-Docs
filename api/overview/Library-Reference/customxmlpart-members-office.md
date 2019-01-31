---
title: CustomXMLPart members (Office)
ms.prod: office
ms.assetid: 76fe85f4-5a35-7d12-2989-6f17a094dcdf
ms.date: 01/30/2019
localization_priority: Normal
---


# CustomXMLPart members (Office)

Represents a single **CustomXMLPart** in a **CustomXMLParts** collection.


## Events

|Name|Description|
|:-----|:-----|
|[NodeAfterDelete](../../Office.CustomXMLPart.NodeAfterDelete.md)|Occurs after a node is deleted in a **CustomXMLPart** object.|
|[NodeAfterInsert](../../Office.CustomXMLPart.NodeAfterInsert.md)|Occurs after a node is inserted in a **CustomXMLPart** object.|
|[NodeAfterReplace](../../Office.CustomXMLPart.NodeAfterReplace.md)|Occurs just after a node is replaced in a **CustomXMLPart** object.|


## Methods

|Name|Description|
|:-----|:-----|
|[AddNode](../../Office.CustomXMLPart.AddNode.md)|Adds a node to the XML tree.|
|[Delete](../../Office.CustomXMLPart.Delete.md)|Deletes the current **CustomXMLPart** from the data store (**IXMLDataStore** interface).|
|[Load](../../Office.CustomXMLPart.Load.md)|Allows the template author to populate a **CustomXMLPart** from an existing file. Returns **True** if the load was successful.|
|[LoadXML](../../Office.CustomXMLPart.LoadXML.md)|Allows the template author to populate a **CustomXMLPart** object from an XML string. Returns **True** if the load was successful.|
|[SelectNodes](../../Office.CustomXMLPart.SelectNodes.md)|Selects a collection of nodes from a custom XML part.|
|[SelectSingleNode](../../Office.CustomXMLPart.SelectSingleNode.md)|Selects a single node within a custom XML part matching an XPath expression.|


## Properties

|Name|Description|
|:-----|:-----|
|[Application](../../Office.CustomXMLPart.Application.md)|Gets an **Application** object that represents the container application for the **CustomXMLPart** object. Read-only.|
|[BuiltIn](../../Office.CustomXMLPart.BuiltIn.md)|Gets a value that indicates whether the **CustomXMLPart** is built-in. Read-only.|
|[Creator](../../Office.CustomXMLPart.Creator.md)|Gets a 32-bit integer that indicates the application in which the **CustomXMLPart** object was created. Read-only.|
|[DocumentElement](../../Office.CustomXMLPart.DocumentElement.md)|Gets the root element of a bound region of data in a document. If the region is empty, the property returns **Nothing**. Read-only.|
|[Errors](../../Office.CustomXMLPart.Errors.md)|Gets a **CustomXMLValidationErrors** object that provides access to any XML validation errors, if any exist. If no validation errors exist, this property returns **Nothing**. Read-only.|
|[Id](../../Office.CustomXMLPart.Id.md)|Gets a **String** containing the GUID assigned to the current **CustomXMLPart** object. Read-only.|
|[NamespaceManager](../../Office.CustomXMLPart.NamespaceManager.md)|Gets the set of namespace prefix mappings used against the current **CustomXMLPart** object. Read-only.|
|[NamespaceURI](../../Office.CustomXMLPart.NamespaceURI.md)|Gets the unique address identifier for the namespace of the **CustomXMLPart** object. Read-only.|
|[Parent](../../Office.CustomXMLPart.Parent.md)|Gets the **Parent** object for the **CustomXMLPart** object. Read-only.|
|[SchemaCollection](../../Office.CustomXMLPart.SchemaCollection.md)|Gets or sets a **CustomXMLSchemaCollection** object representing the set of schemas attached to a bound region of data in a document. Read/write.|
|[XML](../../Office.CustomXMLPart.XML.md)|Gets the XML representation of the current **CustomXMLPart** object. Read-only.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]