Option Explicit

Sub set_header()

With Sheet1.PageSetup
  .LeftHeader = Format(Date, "yyyy/mm/dd")
  .CenterHeader = ""
  .RightHeader = ""
  .LeftFooter = ""
  .CenterFooter = ""
  .RightFooter = ""
End With

End Sub
