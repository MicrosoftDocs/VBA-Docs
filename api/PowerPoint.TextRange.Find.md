---
title: TextRange.Find method (PowerPoint)
keywords: vbapp10.chm569034
f1_keywords:
- vbapp10.chm569034
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.Find
ms.assetid: 24186821-3a0a-efd5-c35a-8b553e00f92b
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRange.Find method (PowerPoint)

Finds the specified text in a text range, and returns a **[TextRange](PowerPoint.TextRange.md)** object that represents the first text range where the text is found. Returns **Nothing** if no match is found.


## Syntax

_expression_.**Find** (_FindWhat_, _After_, _MatchCase_, _WholeWords_)

_expression_ A variable that represents a [TextRange](PowerPoint.TextRange.md) object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FindWhat_|Required|**String**|The text to search for.|
| _After_|Optional|**Long**|The position of the character (in the specified text range) after which you want to search for the next occurrence of FindWhat. For example, if you want to search from the fifth character of the text range, specify 4 for After. If this argument is omitted, the first character of the text range is used as the starting point for the search.|
| _MatchCase_|Optional|**[MsoTriState](Office.MsoTriState.md)**|**msoTrue** for the search to distinguish between uppercase and lowercase characters.|
| _WholeWords_|Optional|**[MsoTriState](Office.MsoTriState.md)**|**msoTrue** for the search to find only whole words and not parts of larger words as well.|

## Return value

TextRange


## Example

This example finds every occurrence of "CompanyX" in the active presentation and formats it as bold.


```vb
For Each sld In Application.ActivePresentation.Slides 
    For Each shp In sld.Shapes 
        If shp.HasTextFrame Then 
            Set txtRng = shp.TextFrame.TextRange 
            Set foundText = txtRng.Find(FindWhat:="CompanyX") 
            Do While Not (foundText Is Nothing) 
                With foundText 
                    .Font.Bold = True 
                    Set foundText = _ 
                        txtRng.Find(FindWhat:="CompanyX", _ 
                        After:=.Start + .Length - 1) 
                End With 
            Loop 
        End If 
    Next 
Next
```


## See also


[TextRange Object](PowerPoint.TextRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
