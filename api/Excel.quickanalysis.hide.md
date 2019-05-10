---
title: QuickAnalysis.Hide method (Excel)
keywords: vbaxl10.chm920074
f1_keywords:
- vbaxl10.chm920074
ms.prod: excel
ms.assetid: dc3b805a-8744-1f63-0509-32b8254958b8
ms.date: 05/10/2019
localization_priority: Normal
---


# QuickAnalysis.Hide method (Excel)

Hides specific members of the Analysis Lens user interface.


## Syntax

_expression_.**Hide** (_XlQuickAnalysisMode_)

_expression_ A variable that represents a **[QuickAnalysis](Excel.quickanalysis.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _XlQuickAnalysisMode_|Optional|**[XlQuickAnalysisMode](excel.xlquickanalysismode.md)**|Indicates for which top level button the callout user interface is displayed. Can be one of the **XlQuickAnalysisMode** constants.|


## Return value

**VOID**


## Remarks

When the argument is set to any one of the following options, the resulting user interface is hidden:

- If missing or set to **0** = Hide all buttons
    
- **1** = If showing, hide the **Conditional Formatting** and **Sparklines** buttons
    
- **2** = If showing, hide the **Charts** button
    
- **3** = If showing, hide the **Suggested Views** button
    
- **4** = If showing, hide the **Totals** button
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]