---
title: Validation.IMEMode property (Excel)
keywords: vbaxl10.chm532076
f1_keywords:
- vbaxl10.chm532076
ms.prod: excel
api_name:
- Excel.Validation.IMEMode
ms.assetid: 0bb1ebc8-257f-b3e0-11d1-b50575e9f86c
ms.date: 05/18/2019
localization_priority: Normal
---


# Validation.IMEMode property (Excel)

Returns or sets the description of the Japanese input rules. Can be one of the **[XlIMEMode](Excel.XlIMEMode.md)** constants listed in the following table. Read/write **Long**.


## Syntax

_expression_.**IMEMode**

_expression_ A variable that represents a **[Validation](Excel.Validation.md)** object.


## Remarks

|Constant|Description|
|:-----|:-----|
| **xlIMEModeAlpha**|Half-width alphanumeric|
| **xlIMEModeAlphaFull**|Full-width alphanumeric|
| **xlIMEModeDisable**|Disable|
| **xlIMEModeHiragana**|Hiragana|
| **xlIMEModeKatakana**|Katakana|
| **xlIMEModeKatakanaHalf**|Katakana (half-width)|
| **xlIMEModeNoControl**|No control|
| **xlIMEModeOff**|Off (English mode)|
| **xlIMEModeOn**|On|

Note that this property can be set only when Japanese language support has been installed and selected.


## Example

This example sets the data input rule for cell E5.

```vb
With Range("E5").Validation 
    .Add Type:=xlValidateWholeNumber, _ 
        AlertStyle:= xlValidAlertStop, _ 
        Operator:=xlBetween, Formula1:="5", Formula2:="10" 
    .InputTitle = "???" 
    .ErrorTitle = "???" 
    .InputMessage = "5??10?????????????" 
    .ErrorMessage = "???????5??10???????" 
    .IMEMode = xlIMEModeAlpha 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]