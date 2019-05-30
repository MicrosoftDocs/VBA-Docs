---
title: Worksheet.XmlMapQuery method (Excel)
keywords: vbaxl10.chm175159
f1_keywords:
- vbaxl10.chm175159
ms.prod: excel
api_name:
- Excel.Worksheet.XmlMapQuery
ms.assetid: ac1d20f4-92ad-110e-00be-0fe4e098cb35
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.XmlMapQuery method (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents the cells mapped to a particular XPath. Returns **Nothing** if the specified XPath has not been mapped to the worksheet.


## Syntax

_expression_.**XmlMapQuery** (_XPath_, _SelectionNamespaces_, _Map_)

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _XPath_|Required| **String**|The XPath to query for.|
| _SelectionNamespaces_|Optional| **Variant**|A space-delimited **String** that contains the namespaces referenced in the XPath parameter. A run-time error is generated if one of the specified namespaces cannot be resolved.|
| _Map_|Optional| **Variant**|Specify an XML map if you want to query for the XPath within a specific map.|

## Return value

**Range**


## Remarks

Unlike the **[XmlDataQuery](Excel.Worksheet.XmlDataQuery.md)** method, the **XmlMapQuery** method returns the entire column of an XML list, including the header row.

> [!NOTE] 
> **XmlMapQuery** allows developers to query for the existence of particular maps. It cannot be used to query for a piece of data in a map. 
> 
> For example, it is valid for a mapped range to exist in which the XPath for that range is `"/root/People[@Age="23"]/FirstName"`. An **XmlMapQuery** call for this XPath returns the correct range. However, a query for `"/root/People[FirstName="Joe"]"` expecting to find "Joe" within the above mapped range fails because the XPath definitions for the mapped ranges are different.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]