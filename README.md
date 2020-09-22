<div align="center">

## Get Recyclebin Info


</div>

### Description

This is not my code. I read it on some website. This code retrieves info about Recyclebin ( number of items in the bin, size of these items).

Here's the website address: http://math.msu.su/~vfnik/WinApi/index.html
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chinese Guy](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chinese-guy.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chinese-guy-get-recyclebin-info__1-45806/archive/master.zip)





### Source Code

<p> I recommend going to the website
"http://math.msu.su/~vfnik/WinApi/index.html"
</p>
<p>throw two textboxes, one commandbutton on the form </p>
<p> give the names txtBinSize & txtNumItems to textboxes </p>
<p>Private Const S_OK = &H0</p>
<p>Private Type ULARGE_INTEGER</p>
<p>   LowPart As Long</p>
<p>   HighPart As Long</p>
<p>End Type</p>
<p>Private Type SHQUERYRBINFO </p>
<p>   cbSize As Long</p>
<p>   i64Size As ULARGE_INTEGER</p>
<p>   i64NumItems As ULARGE_INTEGER</p>
<p>End Type</p>
<p>Private Declare Function SHQueryRecycleBin Lib "shell32.dll" _
    Alias "SHQueryRecycleBinA" (ByVal pszRootPath As String, _
    pSHQueryRBInfo As SHQUERYRBINFO) As Long
</p>
<p>Private Sub Command1_Click()</p>
<p>   ' Display the number of items in the Recycle Bin on the C: drive and the size of it.
</p>
<p> 'information about the bin </p>
  <p>  Dim rbinfo As SHQUERYRBINFO </p>
<p>   Dim retval As Long ' return value </p>
<p>   ' Initialize the size of the structure.</p>
<p>   rbinfo.cbSize = Len(rbinfo)</p>
<p>   ' Query the contents of C:'s Recycle Bin.</p>
<p>   ' the path doesn't have to be the root path</p>
<p>   retval = SHQueryRecycleBin("C:\", rbinfo)</p>
  <p>  ' Display the number of items in the Recycle Bin, if the value is
   ' within Visual Basic's numeric display limits.</p>
<p>   If (rbinfo.i64NumItems.LowPart And &H80000000) = &H80000000 Or _
   rbinfo.i64NumItems.HighPart > 0 Then </p>
<p>  txtNumItems = "Recycle Bin contains more than 2,147,483,647 items." </p>
<p>   Else </p>
<p>      txtNumItems = "Recycle Bin contains " & rbinfo.i64NumItems.LowPart & " items." </p>
<p>   End If </p>
<p>   ' Likewise display the number of bytes the Recycle Bin is taking up.</p>
<p>   If (rbinfo.i64Size.LowPart And &H80000000) = &H80000000 Or rbinfo.i64Size.HighPart > 0 Then </p>
<p>      txtBinSize = "Recycle Bin consumes more than 2,147,483,647 bytes." </p>
<p>   Else </p>
<p>      txtBinSize = "Recycle Bin consumes " & rbinfo.i64Size.LowPart & " bytes." </p>
  <p>  End If</p>
<p>End Sub </p>

