Attribute VB_Name = "modCast"
Private Type UDT_MAKE
 Name As String
 Value As String
 Content(1) As New Collection
 Extra(1) As New Collection
 More(1) As New Collection
End Type

Private Sub AddContent(c() As Collection, ByVal Name As String, ByVal Value As String)
'sub to add or replace properties of a made var
Dim l As Long

For l& = 1 To c(0).Count
 If LCase$(c(0).Item(l&)) = LCase$(Name$) Then
  c(0).Remove l&: c(1).Remove l&
 If l& <> c(0).Count Then
  c(0).Add Name$: c(1).Add Value$
 Else
  c(0).Add Name$, , l&: c(1).Add Value$, , l&
 End If
  Exit Sub
 End If
Next l&

c(0).Add Name$: c(1).Add Value$
End Sub

Public Function Execute(ByVal sData As String) As String
Dim lA&, lB&, lC&, l&, k&
Dim make() As UDT_MAKE
Dim arr() As String
Dim arrC() As String
Dim v As Variant, c As Variant

lA& = InStr(LCase$(sData$), "<script ")
lB& = InStr(lA& + 1, LCase$(sData$), "language=""castscript""")
lA& = InStr(lB& + 1, LCase$(sData$), ">")

lB& = InStr(lA& + 1, LCase$(sData$), "</script>")

arr$() = newSplit(Mid$(sData$, lA& + 1, lB& - lA& - 1), vbCrLf)

For Each v In arr$()
s$ = newTrim$(v)
 If s$ <> "" Then
  lC& = InStr(s$, " ")
  Select Case Trim$(LCase$(Mid$(s$, 1, lC& - 1)))
   Case "make"
     ReDim Preserve make(l&)

    s$ = Trim$(Mid$(s$, lC& + 1))
     lA& = InStr(s$, ":")

     make(l&).Name = Trim$(Mid$(s$, 1, lA& - 1))
     make(l&).Value = Trim$(Mid$(s$, lA& + 1))
    l& = l& + 1
   Case "set"
    ReDim Preserve make(l&)
    s$ = Trim$(Mid$(s$, lC& + 1))
    
    lA& = InStr(s$, ":")

      If InStr(LCase$(s$), " but ") <> 0 And InStr(LCase$(s$), " but ") < InStr(LCase$(s$), "{") Then
       k& = GetIndex&(make(), Trim$(Mid$(s$, lA& + 1, InStr(s$, " BUT ") - lA& - 1)))
        If k& = -1 Then GoTo 1
        Call SetNew(make(l&), make(k&))
       
        make(l&).Name = Trim$(Mid$(s$, 1, lA& - 1))

         k& = GetIndex&(make(), Trim$(LCase$(Mid$(s$, 1, lA& - 1))))
          If k& = -1 Then GoTo 1
      s$ = Trim$(Mid$(s$, InStr(s$, " BUT ") + 6))
       lA& = InStr(s$, "{")
       lB& = InStr(lA& + 1, s$, "}")
      s$ = Trim$(Mid$(s$, lA& + 1, lB& - lA& - 1))

      arrC$() = newSplit(s$, ";")
        i = make(l&).Content(0).Count & make(l&).Content(1).Count
       For Each c In arrC$()
        s$ = newTrim$(c)
          If s$ <> "" Then
           lA& = InStr(s$, "=")
            Select Case Trim$(LCase$(Mid$(s$, 1, lA& - 1)))
           Case "extra"
            Call ParseExtra(make(l&), Mid$(s$, lA& + 1))
           Case "before"
            Call ParseMore(make(l&), Mid$(s$, lA& + 1))
           Case "after"
            Call ParseMore(make(l&), Mid$(s$, lA& + 1), True)
           Case Else
            Call AddContent(make(l&).Content(), Mid$(s$, 1, lA& - 1), Mid$(s$, lA& + 1))
          End Select
         End If
       Next c
      Else
       k& = GetIndex&(make(), Mid$(s$, lA& + 1))
        If k& = -1 Then GoTo 1
       make(l&) = make(k&)
       make(l&).Name = Mid$(s$, 1, lA& - 1)
      End If
     l& = l& + 1
1
   Case Else
    k& = GetIndex&(make(), LCase$(Mid$(s$, 1, lC& - 1)))
     If k& = -1 Then GoTo 2
      s$ = Mid$(s$, lC& + 1)
      lA& = InStr(s$, "{")
      lB& = InStr(lA& + 1, s$, "}")
      s$ = Mid$(s$, lA& + 1, lB& - lA& - 1)

      arrC$() = newSplit(s$, ";")
      i = make(k&).Content(0).Count & make(k&).Content(1).Count
       For Each c In arrC$()
        s$ = newTrim$(c)
         If s$ <> "" Then
          lA& = InStr(s$, "=")
          Select Case Trim$(LCase$(Mid$(s$, 1, lA& - 1)))
           Case "extra"
            Call ParseExtra(make(k&), Mid$(s$, lA& + 1))
           Case "before"
            Call ParseMore(make(k&), Mid$(s$, lA& + 1))
           Case "after"
            Call ParseMore(make(k&), Mid$(s$, lA& + 1), True)
           Case Else
            Call AddContent(make(k&).Content(), Mid$(s$, 1, lA& - 1), Mid$(s$, lA& + 1))
          End Select
         End If
       Next c
2
  End Select
 End If
Next v

s$ = sData$
For k& = 0 To l& - 1

s$ = ReplaceTags(s$, make(k&).Name, make(k&).Value, RetContent$(make(k&).Content()), RetMore$(make(k&).More()), RetExtra$(make(k&).Extra()), RetExtra$(make(k&).Extra(), True), RetMore$(make(k&).More(), True))

 's$ = Replace(s$, "<" & make(k&).Name & ">", RetMore$(make(k&).More()) & "<" & make(k&).Value & " " & RetContent$(make(k&).Content()) & ">" & RetExtra$(make(k&).Extra()), , , vbTextCompare)
 's$ = Replace$(s$, "</" & make(k&).Name & ">", RetExtra$(make(k&).Extra(), True) & "</" & make(k&).Value & ">" & RetMore$(make(k&).More(), True), , , vbTextCompare)
Next k&
Execute$ = s$
End Function

Private Function GetIndex(mk() As UDT_MAKE, ByVal Name As String) As Long
'returns the index of a mad var or -1 if no var found
Dim l As Long

For l& = 0 To UBound(mk)
 If LCase$(mk(l&).Name) = LCase$(Name$) Then GetIndex& = l&: Exit Function
Next l&
GetIndex& = -1
End Function

Private Function newSplit(ByVal sTxt As String, ByVal Del As String)
'new split function that ignores any instance of the delimiter
'in between {}'s
Dim Data() As String, i As Integer, k As Integer
Dim sA As String, t As Integer, c As Integer

  k% = 0: ReDim Data(k%) As String
c% = 0
For i% = 1 To Len(sTxt$)

  sA$ = Mid$(sTxt$, i%, 1)

  If sA$ = Chr(34) Then
   If c% = 0 Then c% = 1 Else c% = 0
  End If

  If sA$ = "{" Then
    t% = t% + 1
  ElseIf sA$ = "}" Then
    t% = t% - 1
    If t% < 0 Then t% = 0
  End If
  
    If Mid$(sTxt$, i%, Len(Del$)) = Del$ And t% = 0 And c% = 0 Then
     i% = i% + Len(Del$) - 1
     k% = k% + 1
     ReDim Preserve Data(k%) As String
    Else
     Data$(k%) = Data$(k%) & sA$
    End If

Next i%

newSplit = Data$
End Function

Private Function newTrim(ByVal sTxt As String) As String
'trims space's,tab's and comments
On Error GoTo 1
Dim sA As String, i As Integer, k As Integer, arr() As String
sA$ = Trim$(sTxt$)

    If sA$ = "" Then newTrim$ = "": Exit Function

i% = InStrRev(sA$, "#") ' find comment marker

If i% <> 0 Then
 arr$() = newSplit(sA$, "#")
 sA$ = Trim$(arr$(0)) ' delete comments
End If

    If sA$ = "" Then newTrim$ = "": Exit Function

For i% = 1 To Len(sA) ' loop thru and strip TAB chr from start to end
  If Mid$(sA$, i%, 1) <> Chr(9) Then i% = i% - 1: Exit For
Next i%

  sA$ = Trim$(Right$(sA$, Len(sA$) - i%)) ' strip tab

    If sA$ = "" Then newTrim$ = "": Exit Function

For i% = Len(sA) To 1 Step -1 ' loop thru and strip TAB chr from end to start
  If Mid$(sA$, i%, 1) <> Chr(9) Then i% = i% + 1: Exit For
Next i%

  sA$ = Trim$(Left$(sA$, i% - 1)) ' strip tab
1
   newTrim$ = sA$
End Function

Private Function ParseExtra(mk As UDT_MAKE, ByVal Values As String)
Dim arrX() As String
Dim x As Variant

arrX$() = newSplit(Values, ",")

For Each x In arrX$()

 If Left$(x, 1) = """" And Right$(x, 1) = """" Then
  mk.Extra(0).Add Mid$(x, 2, Len(x) - 2)
  mk.Extra(1).Add ""
 Else
  mk.Extra(0).Add "<" & x & ">"
  If InStr(x, " ") Then x = Trim$(Mid$(x, 1, InStr(x, " ") - 1))
  mk.Extra(1).Add "</" & x & ">"
 End If
Next x
End Function

Private Function ParseMore(mk As UDT_MAKE, ByVal Values As String, Optional ByVal After As Boolean = False)
Dim arrX() As String
Dim x As Variant, s As String

arrX$() = newSplit(Values, ",")

For Each x In arrX$()
 If Left$(x, 1) = """" And Right$(x, 1) = """" Then s$ = Mid$(x, 2, Len(x) - 2) _
 Else s$ = "<" & x & ">"

 If After = False Then
  mk.More(0).Add s$
 Else
  mk.More(1).Add s$
 End If
Next x
End Function

Private Function Insert(ByVal outStr As String, ByVal insStr As String, ByVal lStart As Long, ByVal lLen As Long, Optional ByVal Retain As Boolean) As String
If Retain = False Then
 Insert$ = Mid$(outStr, 1, lStart& - 1) & insStr$ & Mid$(outStr$, lStart& + lLen&)
Else
 Insert$ = Mid$(outStr, 1, lStart& - 1) & insStr$ & Mid$(outStr$, lStart&)
End If
End Function

Private Function ReplaceTags(ByVal Str As String, ByVal TagName As String, ByVal Value As String, ByVal Content As String, ByVal Before As String, ByVal Extra As String, ByVal CExtra As String, ByVal After As String) As String
On Error GoTo 1
Dim sTag As String
Dim go As Boolean

  Str$ = Replace(Str$, "</" & TagName$ & ">", CExtra$ & Before$ & "</" & Value$ & ">" & After$, , , vbTextCompare)

 a& = InStr(Str$, "<")
 b& = InStr(a& + 1, Str$, " ")
 c& = InStr(a& + 1, Str$, ">")
 d& = InStr(c& + 1, Str$, "<")

Do

  If b& >= a& And b& <= c& Then
   sTag$ = Mid$(Str$, a& + 1, b& - a& - 1)
  Else
   sTag$ = Mid$(Str$, a& + 1, c& - a& - 1)
  End If

 If Left$(sTag$, 1) <> "/" Then
  If LCase$(sTag$) = LCase$(TagName$) Then
   If b& >= a& And b& <= c& Then
    Str$ = Insert$(Str$, Extra$, c& + 1, d& - c& - 1, True)
    Str$ = Insert$(Str$, Value$ & " " & Content$, a& + 1, b& - a& - 1)
   Else
    Str$ = Insert$(Str$, ">" & Extra$, c& + 1, d& - c&, True)
    Str$ = Insert$(Str$, Value$ & " " & Content$, a& + 1, c& - a&)
    sMore$ = ""
   End If
  End If

  'These lines are to be used if you want to limit the script to tags after the BODY tag
  'If go = True Then
  'End If
  'If LCase$(sTag$) = "body" Then go = True
 End If

 a& = InStr(c& + 1, Str$, "<")
 b& = InStr(a& + 1, Str$, " ")
 c& = InStr(a& + 1, Str$, ">")
 d& = InStr(c& + 1, Str$, "<")
 
 If a& = 0 Then Exit Do
 DoEvents
Loop
ReplaceTags$ = Str$
Exit Function
1
ReplaceTags$ = "Compile Error"
End Function

Private Function RetContent(c() As Collection) As String
On Error GoTo 1
Dim l As Long
Dim s As String

For l& = 1 To c(0).Count
 s$ = s$ & c(0).Item(l&) & "=" & c(1).Item(l&) & " "
Next l&
1
RetContent$ = Trim$(s$)
End Function

Private Function RetExtra(c() As Collection, Optional ByVal CloseTag As Boolean = False) As String
On Error GoTo 1
Dim l As Long
Dim s As String

If CloseTag = True Then
 If c(1).Count > 0 Then
  For l& = c(1).Count To 1 Step -1
   s$ = s$ & c(1).Item(l&)
  Next l&
 End If
Else
 If c(0).Count > 0 Then
  For l& = 1 To c(0).Count
   s$ = s$ & c(0).Item(l&)
  Next l&
 End If
End If
1
RetExtra$ = s$
End Function

Private Function RetMore(c() As Collection, Optional ByVal After As Boolean = False) As String
On Error GoTo 1
Dim l As Long
Dim s As String

If After = True Then
 If c(1).Count > 0 Then
  For l& = c(1).Count To 1 Step -1
   s$ = s$ & c(1).Item(l&)
  Next l&
 End If
Else
 If c(0).Count > 0 Then
  For l& = 1 To c(0).Count
   s$ = s$ & c(0).Item(l&)
  Next l&
 End If
End If
1
RetMore$ = s$
End Function

Private Sub SetNew(mk As UDT_MAKE, mk2 As UDT_MAKE)
mk.Name = mk2.Name
mk.Value = mk2.Value

If mk2.Content(0).Count > 0 Then
 For l& = 1 To mk2.Content(0).Count
  mk.Content(0).Add mk2.Content(0).Item(l&)
  mk.Content(1).Add mk2.Content(1).Item(l&)
 Next l&
End If

If mk2.Extra(0).Count > 0 Then
 For l& = 1 To mk2.Extra(0).Count
  mk.Extra(0).Add mk2.Extra(0).Item(l&)
  mk.Extra(1).Add mk2.Extra(1).Item(l&)
 Next l&
End If

If mk2.More(0).Count > 0 Then
 For l& = 1 To mk2.More(0).Count
  mk.More(0).Add mk2.More(0).Item(l&)
 Next l&
End If

If mk2.More(1).Count > 0 Then
 For l& = 1 To mk2.More(1).Count
  mk.More(1).Add mk2.More(1).Item(l&)
 Next l&
End If
End Sub
