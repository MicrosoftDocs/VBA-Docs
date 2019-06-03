---
title: ComboBox object (Access)
keywords: vbaac10.chm11545
f1_keywords:
- vbaac10.chm11545
ms.prod: access
api_name:
- Access.ComboBox
ms.assetid: 1cf508d5-023e-eb38-3991-71e82b2a4e7e
ms.date: 02/27/2019
localization_priority: Normal
---


# ComboBox object (Access)

This object corresponds to a combo box control. The combo box control combines the features of a text box and a list box. Use a combo box when you want the option of either typing a value or selecting a value from a predefined list.


## Remarks

|Control|Tool|
|:-----|:-----|
|![Combo box control](../images/t-combox_ZA06053980.gif)|![Combo box tool](../images/a_combobox_ZA06047114.gif)|

In Form view, Microsoft Access doesn't display the list until you click the combo box's arrow.

If you have Control Wizards on before you select the combo box tool, you can create a combo box with a wizard. To turn Control Wizards on or off, click the **Control Wizards** tool in the toolbox.

The setting of the **LimitToList** property determines whether you can enter values that aren't in the list.

The list can be single- or multiple-column, and the columns can appear with or without headings.
    

## Example

The following example shows how to use multiple **ComboBox** controls to supply criteria for a query.

```vb
Private Sub cmdSearch_Click()
    Dim db As Database
    Dim qd As QueryDef
    Dim vWhere As Variant
    
    Set db = CurrentDb()
    
    On Error Resume Next
    db.QueryDefs.Delete "Query1"
    On Error GoTo 0
    
    vWhere = Null
    vWhere = vWhere & " AND [PymtTypeID]=" & Me.cboPaymentTypes
    vWhere = vWhere & " AND [RefundTypeID]=" & Me.cboRefundType
    vWhere = vWhere & " AND [RefundCDMID]=" & Me.cboRefundCDM
    vWhere = vWhere & " AND [RefundOptionID]=" & Me.cboRefundOption
    vWhere = vWhere & " AND [RefundCodeID]=" & Me.cboRefundCode
    
    If Nz(vWhere, "") = "" Then
        MsgBox "There are no search criteria selected." & vbCrLf & vbCrLf & _
        "Search Cancelled.", vbInformation, "Search Canceled."
        
    Else
        Set qd = db.CreateQueryDef("Query1", "SELECT * FROM tblRefundData WHERE " & _
        Mid(vWhere, 6))
        
        db.Close
        Set db = Nothing
        
        DoCmd.OpenQuery "Query1", acViewNormal, acReadOnly
    End If
End Sub
```

<br/>

The following example shows how to set the **RowSource** property of a combo box when a form is loaded. When the form is displayed, the items stored in the **Departments** field of the **tblDepartment** combo box are displayed in the **cboDept** combo box.

```vb
Private Sub Form_Load()
    Me.Caption = "Today is " & Format$(Date, "dddd mmm-d-yyyy")
    Me.RecordSource = "tblDepartments"
    DoCmd.Maximize  
    txtDept.ControlSource = "Department"
    cmdClose.Caption = "&Close"
    cboDept.RowSourceType = "Table/Query"
    cboDept.RowSource = "SELECT Department FROM tblDepartments"
End Sub
```

<br/>

The following example shows how to create a combo box that is bound to one column while displaying another. Setting the **ColumnCount** property to 2 specifies that the **cboDept** combo box will display the first two columns of the data source specified by the **RowSource** property. Setting the **BoundColumn** property to 1 specifies that the value stored in the first column will be returned when you inspect the value of the combo box.

The **ColumnWidths** property specifies the width of the two columns. By setting the width of the first column to **0in.**, the first column is not displayed in the combo box.

```vb
Private Sub cboDept_Enter()
    With cboDept
        .RowSource = "SELECT * FROM tblDepartments ORDER BY Department"
        .ColumnCount = 2
        .BoundColumn = 1
        .ColumnWidths = "0in.;1in."
    End With
End Sub
```

<br/>

The following example shows how to add an item to a bound combo box.

```vb
Private Sub cboMainCategory_NotInList(NewData As String, Response As Integer)

    On Error GoTo Error_Handler
    Dim intAnswer As Integer
    intAnswer = MsgBox("""" & NewData & """ is not an approved category. " & vbcrlf _
        & "Do you want to add it now?", vbYesNo + vbQuestion, "Invalid Category")

    Select Case intAnswer
        Case vbYes
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO tlkpCategoryNotInList (Category) " & _ 
                         "Select """ & NewData & """;"
            DoCmd.SetWarnings True
            Response = acDataErrAdded
        Case vbNo
            MsgBox "Please select an item from the list.", _
                vbExclamation + vbOKOnly, "Invalid Entry"
            Response = acDataErrContinue

    End Select

    Exit_Procedure:
        DoCmd.SetWarnings True
        Exit Sub

    Error_Handler:
        MsgBox Err.Number & ", " & Err.Description
        Resume Exit_Procedure
        Resume

End Sub
```

## Events

- [AfterUpdate](Access.ComboBox.AfterUpdate-event.md)
- [BeforeUpdate](Access.ComboBox.BeforeUpdate-event.md)
- [Change](Access.ComboBox.Change.md)
- [Click](Access.ComboBox.Click.md)
- [DblClick](Access.ComboBox.DblClick.md)
- [Dirty](Access.ComboBox.Dirty.md)
- [Enter](Access.ComboBox.Enter.md)
- [Exit](Access.ComboBox.Exit.md)
- [GotFocus](Access.ComboBox.GotFocus.md)
- [KeyDown](Access.ComboBox.KeyDown.md)
- [KeyPress](Access.ComboBox.KeyPress.md)
- [KeyUp](Access.ComboBox.KeyUp.md)
- [LostFocus](Access.ComboBox.LostFocus.md)
- [MouseDown](Access.ComboBox.MouseDown.md)
- [MouseMove](Access.ComboBox.MouseMove.md)
- [MouseUp](Access.ComboBox.MouseUp.md)
- [NotInList](Access.ComboBox.NotInList.md)
- [Undo](Access.ComboBox.Undo(even).md)

## Methods

- [AddItem](Access.ComboBox.AddItem.md)
- [Dropdown](Access.ComboBox.Dropdown.md)
- [Move](Access.ComboBox.Move.md)
- [RemoveItem](Access.ComboBox.RemoveItem.md)
- [Requery](Access.ComboBox.Requery.md)
- [SetFocus](Access.ComboBox.SetFocus.md)
- [SizeToFit](Access.ComboBox.SizeToFit.md)
- [Undo](Access.ComboBox.Undo(method).md)

## Properties

- [AddColon](Access.ComboBox.AddColon.md)
- [AfterUpdate](Access.ComboBox.AfterUpdate-property.md)
- [AllowAutoCorrect](Access.ComboBox.AllowAutoCorrect.md)
- [AllowValueListEdits](Access.ComboBox.AllowValueListEdits.md)
- [Application](Access.ComboBox.Application.md)
- [AutoExpand](Access.ComboBox.AutoExpand.md)
- [AutoLabel](Access.ComboBox.AutoLabel.md)
- [BackColor](Access.ComboBox.BackColor.md)
- [BackShade](Access.ComboBox.BackShade.md)
- [BackStyle](Access.ComboBox.BackStyle.md)
- [BackThemeColorIndex](Access.ComboBox.BackThemeColorIndex.md)
- [BackTint](Access.ComboBox.BackTint.md)
- [BeforeUpdate](Access.ComboBox.BeforeUpdate-property.md)
- [BorderColor](Access.ComboBox.BorderColor.md)
- [BorderShade](Access.ComboBox.BorderShade.md)
- [BorderStyle](Access.ComboBox.BorderStyle.md)
- [BorderThemeColorIndex](Access.ComboBox.BorderThemeColorIndex.md)
- [BorderTint](Access.ComboBox.BorderTint.md)
- [BorderWidth](Access.ComboBox.BorderWidth.md)
- [BottomMargin](Access.ComboBox.BottomMargin.md)
- [BottomPadding](Access.ComboBox.BottomPadding.md)
- [BoundColumn](Access.ComboBox.BoundColumn.md)
- [CanGrow](Access.ComboBox.CanGrow.md)
- [CanShrink](Access.ComboBox.CanShrink.md)
- [Column](Access.ComboBox.Column.md)
- [ColumnCount](Access.ComboBox.ColumnCount.md)
- [ColumnHeads](Access.ComboBox.ColumnHeads.md)
- [ColumnHidden](Access.ComboBox.ColumnHidden.md)
- [ColumnOrder](Access.ComboBox.ColumnOrder.md)
- [ColumnWidth](Access.ComboBox.ColumnWidth.md)
- [ColumnWidths](Access.ComboBox.ColumnWidths.md)
- [Controls](Access.ComboBox.Controls.md)
- [ControlSource](Access.ComboBox.ControlSource.md)
- [ControlTipText](Access.ComboBox.ControlTipText.md)
- [ControlType](Access.ComboBox.ControlType.md)
- [DecimalPlaces](Access.ComboBox.DecimalPlaces.md)
- [DefaultValue](Access.ComboBox.DefaultValue.md)
- [DisplayAsHyperlink](Access.ComboBox.DisplayAsHyperlink.md)
- [DisplayWhen](Access.ComboBox.DisplayWhen.md)
- [Enabled](Access.ComboBox.Enabled.md)
- [EventProcPrefix](Access.ComboBox.EventProcPrefix.md)
- [FontBold](Access.ComboBox.FontBold.md)
- [FontItalic](Access.ComboBox.FontItalic.md)
- [FontName](Access.ComboBox.FontName.md)
- [FontSize](Access.ComboBox.FontSize.md)
- [FontUnderline](Access.ComboBox.FontUnderline.md)
- [FontWeight](Access.ComboBox.FontWeight.md)
- [ForeColor](Access.ComboBox.ForeColor.md)
- [ForeShade](Access.ComboBox.ForeShade.md)
- [ForeThemeColorIndex](Access.ComboBox.ForeThemeColorIndex.md)
- [ForeTint](Access.ComboBox.ForeTint.md)
- [Format](Access.ComboBox.Format.md)
- [FormatConditions](Access.ComboBox.FormatConditions.md)
- [GridlineColor](Access.ComboBox.GridlineColor.md)
- [GridlineShade](Access.ComboBox.GridlineShade.md)
- [GridlineStyleBottom](Access.ComboBox.GridlineStyleBottom.md)
- [GridlineStyleLeft](Access.ComboBox.GridlineStyleLeft.md)
- [GridlineStyleRight](Access.ComboBox.GridlineStyleRight.md)
- [GridlineStyleTop](Access.ComboBox.GridlineStyleTop.md)
- [GridlineThemeColorIndex](Access.ComboBox.GridlineThemeColorIndex.md)
- [GridlineTint](Access.ComboBox.GridlineTint.md)
- [GridlineWidthBottom](Access.ComboBox.GridlineWidthBottom.md)
- [GridlineWidthLeft](Access.ComboBox.GridlineWidthLeft.md)
- [GridlineWidthRight](Access.ComboBox.GridlineWidthRight.md)
- [GridlineWidthTop](Access.ComboBox.GridlineWidthTop.md)
- [Height](Access.ComboBox.Height.md)
- [HelpContextId](Access.ComboBox.HelpContextId.md)
- [HideDuplicates](Access.ComboBox.HideDuplicates.md)
- [HorizontalAnchor](Access.ComboBox.HorizontalAnchor.md)
- [Hyperlink](Access.ComboBox.Hyperlink.md)
- [IMEHold](Access.ComboBox.IMEHold.md)
- [IMEMode](Access.ComboBox.IMEMode.md)
- [IMESentenceMode](Access.ComboBox.IMESentenceMode.md)
- [InheritValueList](Access.ComboBox.InheritValueList.md)
- [InputMask](Access.ComboBox.InputMask.md)
- [InSelection](Access.ComboBox.InSelection.md)
- [IsHyperlink](Access.ComboBox.IsHyperlink.md)
- [IsVisible](Access.ComboBox.IsVisible.md)
- [ItemData](Access.ComboBox.ItemData.md)
- [ItemsSelected](Access.ComboBox.ItemsSelected.md)
- [KeyboardLanguage](Access.ComboBox.KeyboardLanguage.md)
- [LabelAlign](Access.ComboBox.LabelAlign.md)
- [LabelX](Access.ComboBox.LabelX.md)
- [LabelY](Access.ComboBox.LabelY.md)
- [Layout](Access.ComboBox.Layout.md)
- [LayoutID](Access.ComboBox.LayoutID.md)
- [Left](Access.ComboBox.Left.md)
- [LeftMargin](Access.ComboBox.LeftMargin.md)
- [LeftPadding](Access.ComboBox.LeftPadding.md)
- [LimitToList](Access.ComboBox.LimitToList.md)
- [ListCount](Access.ComboBox.ListCount.md)
- [ListIndex](Access.ComboBox.ListIndex.md)
- [ListItemsEditForm](Access.ComboBox.ListItemsEditForm.md)
- [ListRows](Access.ComboBox.ListRows.md)
- [ListWidth](Access.ComboBox.ListWidth.md)
- [Locked](Access.ComboBox.Locked.md)
- [Name](Access.ComboBox.Name.md)
- [NumeralShapes](Access.ComboBox.NumeralShapes.md)
- [OldBorderStyle](Access.ComboBox.OldBorderStyle.md)
- [OldValue](Access.ComboBox.OldValue.md)
- [OnChange](Access.ComboBox.OnChange.md)
- [OnClick](Access.ComboBox.OnClick.md)
- [OnDblClick](Access.ComboBox.OnDblClick.md)
- [OnDirty](Access.ComboBox.OnDirty.md)
- [OnEnter](Access.ComboBox.OnEnter.md)
- [OnExit](Access.ComboBox.OnExit.md)
- [OnGotFocus](Access.ComboBox.OnGotFocus.md)
- [OnKeyDown](Access.ComboBox.OnKeyDown.md)
- [OnKeyPress](Access.ComboBox.OnKeyPress.md)
- [OnKeyUp](Access.ComboBox.OnKeyUp.md)
- [OnLostFocus](Access.ComboBox.OnLostFocus.md)
- [OnMouseDown](Access.ComboBox.OnMouseDown.md)
- [OnMouseMove](Access.ComboBox.OnMouseMove.md)
- [OnMouseUp](Access.ComboBox.OnMouseUp.md)
- [OnNotInList](Access.ComboBox.OnNotInList.md)
- [OnUndo](Access.ComboBox.OnUndo.md)
- [Parent](Access.ComboBox.Parent.md)
- [Properties](Access.ComboBox.Properties.md)
- [ReadingOrder](Access.ComboBox.ReadingOrder.md)
- [Recordset](Access.ComboBox.Recordset.md)
- [RightMargin](Access.ComboBox.RightMargin.md)
- [RightPadding](Access.ComboBox.RightPadding.md)
- [RowSource](Access.ComboBox.RowSource.md)
- [RowSourceType](Access.ComboBox.RowSourceType.md)
- [ScrollBarAlign](Access.ComboBox.ScrollBarAlign.md)
- [Section](Access.ComboBox.Section.md)
- [Selected](Access.ComboBox.Selected.md)
- [SelLength](Access.ComboBox.SelLength.md)
- [SelStart](Access.ComboBox.SelStart.md)
- [SelText](Access.ComboBox.SelText.md)
- [SeparatorCharacters](Access.ComboBox.SeparatorCharacters.md)
- [ShortcutMenuBar](Access.ComboBox.ShortcutMenuBar.md)
- [ShowOnlyRowSourceValues](Access.ComboBox.ShowOnlyRowSourceValues.md)
- [SmartTags](Access.ComboBox.SmartTags.md)
- [SpecialEffect](Access.ComboBox.SpecialEffect.md)
- [StatusBarText](Access.ComboBox.StatusBarText.md)
- [TabIndex](Access.ComboBox.TabIndex.md)
- [TabStop](Access.ComboBox.TabStop.md)
- [Tag](Access.ComboBox.Tag.md)
- [Text](Access.ComboBox.Text.md)
- [TextAlign](Access.ComboBox.TextAlign.md)
- [ThemeFontIndex](Access.ComboBox.ThemeFontIndex.md)
- [Top](Access.ComboBox.Top.md)
- [TopMargin](Access.ComboBox.TopMargin.md)
- [TopPadding](Access.ComboBox.TopPadding.md)
- [ValidationRule](Access.ComboBox.ValidationRule.md)
- [ValidationText](Access.ComboBox.ValidationText.md)
- [Value](Access.ComboBox.Value.md)
- [VerticalAnchor](Access.ComboBox.VerticalAnchor.md)
- [Visible](Access.ComboBox.Visible.md)
- [Width](Access.ComboBox.Width.md)




## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
