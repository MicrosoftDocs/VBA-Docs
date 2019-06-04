---
title: Application.Presentations property (PowerPoint)
keywords: vbapp10.chm502001
f1_keywords:
- vbapp10.chm502001
ms.prod: powerpoint
api_name:
- PowerPoint.Application.Presentations
ms.assetid: d6f5f565-d593-e230-c3b9-2302bdd83644
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Presentations property (PowerPoint)

Returns a  **[Presentations](PowerPoint.Presentations.md)** collection that represents all open presentations. Read-only.


## Syntax

_expression_. `Presentations`

_expression_ A variable that represents an **[Application](PowerPoint.Application.md)** object.


## Return value

Presentations


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../powerpoint/How-to/return-objects-from-collections.md).

If your Visual Studio solution includes the  **Microsoft.Office.Interop.PowerPoint** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.PowerPoint._Application.Presentations**
    

## Example

This example opens the presentation named "Long Version.ppt."


```vb
Application.Presentations.Open _ 
    FileName:="c:\My Documents\Long version.ppt"
```

This example saves presentation one as "Year-End Report.ppt."




```vb
Application.Presentations(1).SaveAs "Year-End Report"
```

This example closes the year-end report presentation.




```vb
Application.Presentations("Year-End Report.ppt").Close
```


## See also


[Application Object](PowerPoint.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]