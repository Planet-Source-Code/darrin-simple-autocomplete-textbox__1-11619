<div align="center">

## Simple AutoComplete TextBox


</div>

### Description

This code will autofill a textbox from a database table using the keyup event.
 
### More Info
 
User input into textbox and database via ADODB recordset.

It will autofill the textbox with the the first match of the letters the user types in.

I used the keyup event to keep it very simple, although if a user types quickly, and presses a key before the previous key is released (keyup'd), it will cause and error and clear the textbox.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Darrin](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/darrin.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/darrin-simple-autocomplete-textbox__1-11619/archive/master.zip)

### API Declarations

```
Place the variables that are in ALL CAPS as global variables, and set STRNAME = "" and the INTPLACE = 0 on the GotFocus and the LostFocus events of the textbox.
You will also need to set "cn" = to an ADODB connection to your database.
```


### Source Code

```
Private Sub Textbox1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim rsTable as ADODB.recordset
Set rsTable = New ADODB.recordset
On Error GoTo ENDOFSUB
 rsTable.Open "Select * from TABLE", cn, adopenstatic, adlockoptomistic
 STRWORD = Me.textbox1.Text
 If Len(STRWORD) < INTPLACE Then
  INTPLACE = Len(STRWORD) - 1
 End If
 If KeyCode = vbKeyBack Or KeyCode = vbKeyLeft Then
  If INTPLACE > 0 Then
   INTPLACE = INTPLACE - 1
   STRWORD = Mid(STRWORD, 1, Len(STRWORD) - 1)
  End If
 ElseIf Me.textbox1.Text = "" Then
  INTPLACE = 0
  STRWORD = ""
 ElseIf KeyCode <> vbKeyDelete And KeyCode <> vbKeyShift Then
  INTPLACE = INTPLACE + 1
  STRWORD = STRWORD & Chr(KeyCode)
 End If
  rsTable.MoveFirst
 If Me.textbox1.Text <> "" Then
  Do While Not rsTable.EOF
    If Mid(Trim(rsTable!Field1), 1, INTPLACE) = UCase(Mid(Me.textbox1.Text, 1, INTPLACE)) Then
     Me.textbox1.Text = Trim(rsTable!Field1)
     Exit Do
    End If
   m_rsEmployee.MoveNext
  Loop
 End If
 If KeyCode <> vbKeyShift Then
  Me.textbox1.SelStart = INTPLACE
  Me.textbox1.SelLength = (Len(Me.textbox1.Text)) - INTPLACE
 End If
 Exit Sub
ENDOFSUB:
 Me.textbox1.Text = ""
 INTPLACE = 0
End Sub
```

