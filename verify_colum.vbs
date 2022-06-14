'*********************************************************
' Objetive: Verify column type in a CSV file
'    @Date: 13/Jun/2022
'  @Author: edcruces99@gmail.com
'
'*********************************************************

'Create Object File
Set FSO = CreateObject("Scripting.FileSystemObject")

'Create Array to check duplicated values
Set arrData =  CreateObject("System.Collections.ArrayList")

'Read First Argument
Filename = WScript.Arguments.Item(0)

'Read First Argument (without extension)
arrFilename=Split(Filename,".")
for each x in arrFilename
    FilenameJustName=x
	Exit For
next

'Result File
Set ff = FSO.OpenTextFile(FilenameJustName & "_result.txt" ,2 , True)

'Open File
Set f1 = fso.OpenTextFile(filename)

'Separator
sep=","

'Counter
tkn_Col1_OK=0
tkn_Col1_Not_OK=0
tkn_Col3_OK=0
tkn_Col3_Not_OK=0
tkn_Col4_OK=0
tkn_Col4_Not_OK=0
tkn_Col5_OK=0
tkn_Col5_Not_OK=0

Do
  linea = f1.ReadLine
  
  arrTkn=Split(linea,sep,-1,vbBinaryCompare) '-1 = regresa todos los tokens

      If IsNumeric(arrTkn(0)) Then 'Col 1 Numerica
	     tkn_Col1_OK=tkn_Col1_OK+1
      Else
	     tkn_Col1_Not_OK=tkn_Col1_Not_OK+1
	  End If
	  '-----------------------------------
      If IsNumeric(arrTkn(2)) Then 'Col 3 Numerica
	     tkn_Col3_OK=tkn_Col3_OK+1
      Else
	     tkn_Col3_Not_OK=tkn_Col3_Not_OK+1
	  End If
	  '-----------------------------------
      If arrTkn(3) = "DATA" Then 'Col 4 Const="DATA"
	     tkn_Col4_OK=tkn_Col4_OK+1
      Else
	     tkn_Col4_Not_OK=tkn_Col4_Not_OK+1
	  End If

      If IsDate(arrTkn(4)) Then 'Col 5 Fecha
	     tkn_Col5_OK=tkn_Col5_OK+1
      Else
	     tkn_Col5_Not_OK=tkn_Col5_Not_OK+1
	  End If
	  
Loop Until f1.AtEndOfStream = true

'Write Result
ff.WriteLine "tkn_Col1_OK: " & tkn_Col1_OK
ff.WriteLine "tkn_Col1_Not_OK: " & tkn_Col1_Not_OK
ff.WriteLine
ff.WriteLine "tkn_Col3_OK: " & tkn_Col3_OK
ff.WriteLine "tkn_Col3_Not_OK: " & tkn_Col3_Not_OK
ff.WriteLine
ff.WriteLine "tkn_Col4_OK: " & tkn_Col4_OK
ff.WriteLine "tkn_Col4_Not_OK: " & tkn_Col4_Not_OK
ff.WriteLine
ff.WriteLine "tkn_Col5_OK: " & tkn_Col5_OK
ff.WriteLine "tkn_Col5_Not_OK: " & tkn_Col5_Not_OK

f1.Close
ff.Close
