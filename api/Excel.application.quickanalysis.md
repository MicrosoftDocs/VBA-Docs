---
title: Application.QuickAnalysis property (Excel)
keywords: vbaxl10.chm133338
f1_keywords:
- vbaxl10.chm133338
ms.prod: excel
ms.assetid: c79c04e7-0caf-470c-ee6d-dc613d6a4cf5
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.QuickAnalysis property (Excel)

Returns a **[QuickAnalysis](Excel.quickanalysis.md)** object that represents the Quick Analysis options of the application.


## Syntax

_expression_.**QuickAnalysis**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Property value

**QUICKANALYSIS**


## Example

The following example displays the Quick Analysis contextual UI with the **Sparklines** option highlighted.

```vb
Sub ShowQuickAnalysisOptions()

'Displays the Quick Analysis contextual UI with the Sparklines option highlighted.
  Application.QuickAnalysis.Show (xlSparklines)

End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]