---
title: Event properties
ROBOTS: INDEX
keywords: vbaac10.chm4998
f1_keywords:
- vbaac10.chm4998
ms.prod: access
ms.assetid: 26c849e7-d433-8d7d-641a-d3171b20d8bd
ms.date: 06/08/2019
localization_priority: Normal
---


# Event properties

**Applies to:** Access 2013 | Access 2016

Event properties cause a macro or the associated Visual Basic event procedure to run when a particular event occurs. For example, if you enter the name of a macro in a command button's **OnClick** property, that macro runs when the command button is clicked.


## Setting

To run a macro, enter the name of the macro. You can choose an existing macro in the list. If the macro is in a macro group, it will be listed under the macro group name, as  _macrogroupname_. _macroname_.

To run the event procedure associated with the event, select **[Event Procedure]** in the list.


> [!NOTE] 
> Although using an event procedure is the recommended method for running Visual Basic code in response to an event, you can also run a user-defined function when an event occurs. To run a user-defined function, place an equal sign (=) before the function name and parentheses after it, as in **=** _functionname_ **( )**.

You can set event properties in the [property sheet](https://msdn.microsoft.com/library/03349d86-f107-9e49-89df-62f55f3a0735%28Office.15%29.aspx) for an object, in a macro, or by using Visual Basic. Note that you can't set any event properties while you are formatting or printing a form or report.

> [!TIP] 
> You can use builders to help you set an event property. To use them, click the **Build** button
![Builder button](../../../images/buildbut_ZA06047218.gif) to the right of the property box, or right-click the property box and then click **Build** on the shortcut menu. In the **Choose Builder** dialog box, select:


- The Macro Builder to create and specify a macro for this event property. You can also use the Macro Builder to edit a macro already specified by the property.
    
- The Code Builder to create and specify an event procedure for this event property. You can also use the Code Builder to edit an event procedure already specified by the property.
    
- In a Microsoft Access database, the Expression Builder to choose and specify a user-defined function for this event property.
    
In Visual Basic, set the property to a string expression.

|**To run this**|**Use this syntax**|**Example**|
|:-----|:-----|:-----|
|Macro|**"** _macroname_ **"**||
|Event procedure|**"[Event Procedure]"**||
|User-defined function|**"=** _functionname_ **( )"**||

## Example

The following example shows how you can use the value entered in the Country control to determine which of two different macros to run when you click the Print Country Report button.


```vb
Private Sub Country_AfterUpdate() 
    If Country = "Canada" Then 
        [Print Country Report].OnClick = "PrintCanadaReport" 
    ElseIf Country = "USA" Then 
        [Print Country Report].OnClick = "PrintUSAReport" 
    End If 
End Sub
```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]