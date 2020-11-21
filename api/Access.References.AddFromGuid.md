---
title: References.AddFromGUID method (Access)
keywords: vbaac10.chm12642
f1_keywords:
- vbaac10.chm12642
ms.prod: access
api_name:
- Access.References.AddFromGuid
ms.assetid: df383ef3-e27c-9590-2ee7-d078060c9313
ms.date: 03/23/2019
localization_priority: Normal
---


# References.AddFromGUID method (Access)

The **AddFromGUID** method creates a **[Reference](Access.Reference.md)** object based on the GUID that identifies a type library. **Reference** object.


## Syntax

_expression_.**AddFromGUID** (_Guid_, _Major_, _Minor_)

_expression_ A variable that represents a **[References](Access.References.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Guid_|Required|**String**|A GUID that identifies a type library.|
| _Major_|Required|**Long**|The major version number of the reference.|
| _Minor_|Required|**Long**|The minor version number of the reference.|

## Return value

Reference


## Remarks

The **[GUID](Access.Reference.Guid.md)** property returns the GUID for a specified **Reference** object. If you stored the value of the **GUID** property, you can use it to re-create a reference that's been broken.

If you add a GUID reference using 0 for both the major and minor version parameters, it will resolve to the latest installed version of an object library.


## Example

The following example re-creates a reference to the **Microsoft Scripting Runtime** version 1.0, based on its GUID on the user's system.

```vb
References.AddFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0

```

The following example adds a reference to the **Microsoft Excel Object Library**, without knowing which version is currently installed.

```vb
References.AddFromGuid "{00020813-0000-0000-C000-000000000046}", 0, 0

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
