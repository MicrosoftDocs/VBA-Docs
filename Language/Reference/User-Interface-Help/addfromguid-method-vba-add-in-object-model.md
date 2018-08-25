---
title: AddFromGuid Method (VBA Add-In Object Model)
keywords: vbob6.chm104045
f1_keywords:
- vbob6.chm104045
ms.prod: office
ms.assetid: d36637d5-6fc6-dbf3-5a2f-7de3c59e8c8c
ms.date: 06/08/2017
---


# AddFromGuid Method (VBA Add-In Object Model)



Adds a reference to the  **References** collection using the globally unique identifier (GUID) of the reference.

## Syntax

_object_**.AddFromGuid(**_guid_, _major_, _minor_**)** **As Reference**
The  **AddFromGuid** syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the Applies To list.|
| _guid_|Required. A [string expression](../../Glossary/vbe-glossary.md#string-expression)representing the GUID of the reference.|
| _major_|Required. A [Long](../../Glossary/vbe-glossary.md) specifying the major version number of the reference.|
| _minor_|Required. A  **Long** specifying the minor version number of the reference.|

## Remarks

The  **AddFromGuid** method searches the[registry](../../Glossary/vbe-glossary.md#registry) to find the reference you want to add. The GUID can be a[type library](../../Glossary/vbe-glossary.md#type-library), [control](../../Glossary/vbe-glossary.md#control), class identifier, and so on.

