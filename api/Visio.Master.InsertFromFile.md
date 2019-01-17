---
title: Master.InsertFromFile Method (Visio)
keywords: vis_sdr.chm10716365
f1_keywords:
- vis_sdr.chm10716365
ms.prod: visio
api_name:
- Visio.Master.InsertFromFile
ms.assetid: 5a24e289-675a-d08b-36f7-0cfaedac5aaf
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.InsertFromFile Method (Visio)

Adds a linked or embedded object to a page, master, or group.


## Syntax

 _expression_. `InsertFromFile`( `_FileName_` , `_Flags_` )

 _expression_ A variable that represents a [Master](./Visio.Master.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The name of the file that contains the object to link or embed.|
| _Flags_|Required| **Integer**|Flags that influence how the object is inserted.|

## Return value

Shape


## Remarks

The  **InsertFromFile** method creates a new shape that represents a linked or embedded OLE object.

The  _Flags_ argument is a bitmask that should be a combination of the following values.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visInsertLink**|&H8|If set, the new shape represents an OLE link to the named file. Otherwise, the  **InsertFromFile** method produces an OLE object from the contents of the named file and embeds it in the document that contains the page, master, or group.|
| **visInsertIcon**|&H10|Displays the new shape as an icon.|

> [!CAUTION] 
> Use caution when you are adding ActiveX controls to your application. ActiveX controls may be designed in such a way that their use could pose a security risk. We recommend that you use controls from trusted sources only.


