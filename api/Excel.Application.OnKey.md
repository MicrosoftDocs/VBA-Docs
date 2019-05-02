---
title: Application.OnKey method (Excel)
keywords: vbaxl10.chm133180
f1_keywords:
- vbaxl10.chm133180
ms.prod: excel
api_name:
- Excel.Application.OnKey
ms.assetid: 43662d8b-19e2-2b4a-4c3a-c64be4007643
ms.date: 04/30/2019
localization_priority: Normal
---


# Application.OnKey method (Excel)

Runs a specified procedure when a particular key or key combination is pressed.


## Syntax

_expression_.**OnKey** (_Key_, _Procedure_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Key_|Required| **String**|A string indicating the key to be pressed.|
| _Procedure_|Optional| **Variant**|A string indicating the name of the procedure to be run. If _Procedure_ is "" (empty text), nothing happens when _Key_ is pressed. This form of **OnKey** changes the normal result of keystrokes in Microsoft Excel.<br/><br/>If _Procedure_ is omitted, _Key_ reverts to its normal result in Microsoft Excel, and any special key assignments made with previous **OnKey** methods are cleared.|

## Remarks

The _Key_ argument can specify any single key combined with Alt, Ctrl, or Shift, or any combination of these keys. Each key is represented by one or more characters, such as `a` for the character a, or `{ENTER}` for the Enter key.

To specify characters that aren't displayed when you press the corresponding key (for example: Enter or Tab), use the codes listed in the following table. Each code in the table represents one key on the keyboard.

|Key|Code|
|:-----|:-----|
|BACKSPACE| `{BACKSPACE}` or `{BS}`|
|BREAK| `{BREAK}`|
|CAPS LOCK| `{CAPSLOCK}`|
|CLEAR| `{CLEAR}`|
|DELETE or DEL| `{DELETE}` or `{DEL}`|
|DOWN ARROW| `{DOWN}`|
|END| `{END}`|
|ENTER (numeric keypad)| `{ENTER}`|
|ENTER| `~` (tilde)|
|ESC|{ `ESCAPE}` or `{ESC}`|
|HELP| `{HELP}`|
|HOME| `{HOME}`|
|INS| `{INSERT}`|
|LEFT ARROW| `{LEFT}`|
|NUM LOCK| `{NUMLOCK}`|
|PAGE DOWN| `{PGDN}`|
|PAGE UP| `{PGUP}`|
|RETURN| `{RETURN}`|
|RIGHT ARROW| `{RIGHT}`|
|SCROLL LOCK| `{SCROLLLOCK}`|
|TAB| `{TAB}`|
|UP ARROW| `{UP}`|
|F1 through F15| `{F1}` through `{F15}`|

<br/>

You can also specify keys combined with Shift and/or Ctrl and/or Alt and/or Command. To specify a key combined with another key or keys, use the following table.

|To combine keys with|Precede the key code by|
|:-----|:-----|
|Shift| `+` (plus sign)|
|Ctrl| `^` (caret)|
|Alt| `%` (percent sign)|
|Command|`*` (asterisk) Only applies to Mac; may only work on Excel 2011 for Mac and not later versions.|

To assign a procedure to one of the special characters (+, ^, %, and so on), enclose the character in braces. For details, see the example.

> [!NOTE] 
> There is no way to currently detect the Command key in recent versions of Office VBA. Microsoft is aware of this and is looking into it.

## Example

This example assigns InsertProc to the key sequence Ctrl+Plus Sign, and assigns SpecialPrintProc to the key sequence Shift+Ctrl+Right Arrow.

```vb
Application.OnKey "^{+}", "InsertProc" 
Application.OnKey "+^{RIGHT}", "SpecialPrintProc"
```

<br/>

This example returns Shift+Ctrl+Right Arrow to its normal meaning.

```vb
Application.OnKey "+^{RIGHT}"
```

<br/>

This example disables the Shift+Ctrl+Right Arrow key sequence.

```vb
Application.OnKey "+^{RIGHT}", ""
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
