---
title: Section.OnFormat property (Access)
keywords: vbaac10.chm12204,vbaac10.chm4089
f1_keywords:
- vbaac10.chm12204,vbaac10.chm4089
ms.prod: access
api_name:
- Access.Section.OnFormat
ms.assetid: 061652a9-0253-8dc2-a8c0-02daa40d132d
ms.date: 03/23/2019
localization_priority: Normal
---


# Section.OnFormat property (Access)

Sets or returns the value of the **On Format** box in the Properties window of a report section. Read/write **String**.


## Syntax

_expression_.**OnFormat**

_expression_ A variable that represents a **[Section](Access.Section.md)** object.


## Remarks

This property is helpful for programmatically changing the action that Microsoft Access takes when an event is triggered. For example, between event calls you may want to change an expression's parameters, or switch from an event procedure to an expression or macro, depending on the circumstances under which the event was triggered. 

The **Format** event occurs when Access determines which data belongs in a report section, but before Access formats the section for previewing or printing.

The **OnFormat** value will be one of the following, depending on the selection chosen in the Choose Builder window (accessed by choosing the **Build** button next to the **On Format** box in the report section's Properties window):

- If you choose Expression Builder, the value will be =_expression_, where _expression_ is the expression from the Expression Builder window.
    
- If you choose Macro Builder, the value is the name of the macro. 
    
- If you choose Code Builder, the value will be [Event Procedure]. 
    
If the **On Format** box is blank, the property value is an empty string.


## Example

The following example prints the value of the **OnFormat** property in the Immediate window for the "GroupHeader0" section in the **Purchase Order** report.

```vb
Debug.Print Reports("Purchase Order").Section("GroupHeader0").OnFormat
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]