---
title: Research.Query method (Excel)
keywords: vbaxl10.chm849073
f1_keywords:
- vbaxl10.chm849073
api_name:
- Excel.Research.Query
ms.assetid: ea3b90ba-9cb4-2682-e092-6e3dd7d40aaf
ms.date: 05/11/2019
ms.localizationpriority: medium
---


# Research.Query method (Excel)

Specifies a research query.


## Syntax

_expression_.**Query** (_ServiceID_, _QueryString_, _QueryLanguage_, _UseSelection_, _RequeryContextXML_, _NewQueryContextXML_, _LaunchQuery_)

_expression_ A variable that represents a **[Research](Excel.Research.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ServiceID_|Required| **String**|Specifies a GUID that identifies the research service.|
| _QueryString_|Optional| **Variant**|Specifies the query string.|
| _QueryLanguage_|Optional| **Variant**|Specifies the query language of the query string.|
| _UseSelection_|Optional| **Boolean**| **True** to use the current selection as the query string. This overrides the _QueryString_ parameter if set. Default value is **False**.|
| _RequeryContextXML_|Optional| **Variant**|Specifies the XML file containing the requested content.|
| _NewQueryContextXML_|Optional| **Variant**|Specifies the XML file containing the new query content.|
| _LaunchQuery_|Optional| **Boolean**| **True** launches the query. **False** displays the Research task pane scoped to search the specified research service.|

## Return value

Variant




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]