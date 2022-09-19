Option Explicit

Sub set_print_area()
    
Dim wksSheet as Worksheet
Set wksSheet = Sheet1
    
wksSheet.PageSetup.PrintArea = wksSheet.Range("A1").CurrentRegion.Address

End Sub
