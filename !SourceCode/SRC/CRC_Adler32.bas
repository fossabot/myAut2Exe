Attribute VB_Name = "CRC_Adler32"
Public Function ADLER32$(Data As StringReader)
   With Data
'            Dim a
            
            Dim L&, H&
            H = 0: L = 1
'            a = GetTickCount
' taken out for performance reason
'               .EOS = False
'               .DisableAutoMove = False
'               Do Until .EOS
'                 'The largest prime less than 2^16
'                  l = (.int8 + l) Mod 65521 '&HFFF1
'                  H = (H + l) Mod 65521 '&HFFF1
'                  If (l And 8) Then myDoEvents
'               Loop
'
'            Debug.Print "a: ", GetTickCount - a 'Benchmark: 20203

 '           a = GetTickCount
               
               Dim StrCharPos&, tmpBuff$
               tmpBuff = StrConv(.mvardata, vbFromUnicode, LocaleID)
'               tmpBuff = .mvardata
               For StrCharPos = 1 To Len(.mvardata)
                  'The largest prime less than 2^16
                  L = (AscB(MidB$(tmpBuff, StrCharPos, 1)) + L) Mod 65521 '&HFFF1
                  H = (H + L) Mod 65521 '&HFFF1
                  
                  If 0 = (StrCharPos Mod &H8000) Then myDoEvents

               Next
'            Debug.Print "b: ", GetTickCount - a 'Benchmark: 5969

      ADLER32 = H16(H) & H16(L)
   End With
End Function
