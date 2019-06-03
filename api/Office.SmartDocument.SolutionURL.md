---
title: SmartDocument.SolutionURL property (Office)
keywords: vbaof11.chm262002
f1_keywords:
- vbaof11.chm262002
ms.prod: office
api_name:
- Office.SmartDocument.SolutionURL
ms.assetid: f4e8af50-9c14-bcc1-ef61-9af3a7c8c65d
ms.date: 01/25/2019
localization_priority: Normal
---


# SmartDocument.SolutionURL property (Office)

Gets or sets an absolute URL that provides the complete path to the XML expansion pack file attached to the active document in Microsoft Word or a workbook in Microsoft Excel. Read/write.


## Syntax

_expression_.**SolutionURL**

_expression_ A variable that represents a **[SmartDocument](Office.SmartDocument.md)** object.


## Remarks

The **SolutionUrl** property returns an empty string when no XML expansion pack is attached to the active document.

Provide appropriate values for the **SolutionID** and **SolutionUrl** properties to attach an available XML expansion pack to the active document and transform it into a smart document without using the **PickSolution** method. 

Set the **SolutionID** and **SolutionUrl** properties to empty strings to remove the attached XML expansion pack.


## Example

The following example determines whether an XML expansion pack is attached to the active Word document, and then displays the smart document's **SolutionURL**.


```vb
 Dim objSmartDoc As Office.SmartDocument 
 Set objSmartDoc = ActiveDocument.SmartDocument 
 If objSmartDoc.SolutionID = "None" Or objSmartDoc.SolutionID = "" Then 
 MsgBox "No XML expansion pack attached." 
 Else 
 MsgBox "Smart document Solution URL: " & _ 
 objSmartDoc.SolutionURL 
 End If 
 Set objSmartDoc = Nothing
```


## See also

- [SmartDocument object members](overview/Library-Reference/smartdocument-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]