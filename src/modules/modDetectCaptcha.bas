Attribute VB_Name = "modDetectCaptcha"
Option Explicit


Function detectSVGCaptcha(ByVal svgCaptcha$) As String
  Dim s$, p, i%, j%, z, k%, re, ms, m, v$, d
  Set d = glbDict: Set re = glbRegex
  re.Pattern = "([MQZ])([^MQZ]*)"
  p = Split(svgCaptcha, " d=\""")
  ms = ListPathAllKeywords
  For i = 0 To UBound(ms)
    If ms(i) <> Empty Then d(ms(i)) = i
  Next
  ReDim z(1 To 6)
  For i = 1 To UBound(p)
    s = Split(p(i), "\""")(0)
    Set ms = re.Execute(s): 'm0 = 0: m1 = 0
    s = re.Replace(s, "$1") 'If Len(s) > 5
    If d.Exists(s) Then
      k = k + 1: z(k) = Array(Val(ms(0).SubMatches(1)), IIf(d(s) <= 26, Chr(d(s) + 65), d(s) - 26))
    End If
  Next
  If k > 0 Then
    For i = 1 To k
      For j = i + 1 To k
        If z(i)(0) > z(j)(0) Then m = z(i): z(i) = z(j): z(j) = m
      Next
    Next
    s = ""
    For i = 1 To k
      s = s & z(i)(1)
    Next
    detectSVGCaptcha = s
  End If
End Function
Private Function glbRegex(Optional bGlobal As Boolean = True, Optional IgnoreCase As Boolean = True, Optional MultiLine As Boolean = True) As Object
  Set glbRegex = CreateObject("VBScript.RegExp")
  With glbRegex: .Global = bGlobal: .IgnoreCase = IgnoreCase: .MultiLine = MultiLine: End With
End Function
Private Function glbDict(Optional ByVal CompareMode As Boolean) As Object
  Set glbDict = CreateObject("Scripting.Dictionary")
  glbDict.CompareMode = -CompareMode
End Function
Private Function ListPathAllKeywords()
  Dim a$, b$, c$, d$, e$, f$, g$, h$, i$, j$, k$, l$, m$, n$, o$, p$, q$, r$, s$, t$, u$, v$, w$, X$, Y$, z$
  Dim k0$, k1$, k2$, k3$, k4$, k5$, k6$, k7$, k8$, k9$, aa
  a = "MQQQQQZMQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQZMQQZ"
  b = "MQQQQQQQQQZMQQQQQQZMQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQZMQQQQQQQQZMQQQQQQQQZ"
  c = "MQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQZ"
  d = "MQQQQQQQQZMQQQQQQQQQQZMQQQQQQQQQQQQQQQZMQQQQQQQZ"
  e = "MQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQZ"
  f = "MQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQZ"
  g = "MQQQQQQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQZ"
  h = "MQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQZ"
  i = ""
  j = "MQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQZ"
  k = "MQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQZ"
  l = ""
  m = "MQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQZ"
  n = "MQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQZ"
  o = ""
  p = "MQQQQQQZMQQQQQQQQQQZMQQQQQQQQQQQQQQQZMQQQQQQQQZ"
  q = "MQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQZ"
  r = "MQQQQQQZMQQQQQQQQQQQQZMQQQQQQQQQQQQQQQZMQQQQQQQQZ"
  s = "MQQQQQQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQZ"
  t = "MQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQZ"
  u = ""
  v = "MQQQQQQQQQQZMQQQQQQQQQQQQQQQQZ"
  w = "MQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQZ"
  X = "MQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQZ"
  Y = "MQQQQQQQQQZMQQQQQQQQQQQQQZ"
  z = "MQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQZ"
  k0 = ""
  k1 = ""
  k2 = "MQQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQQQZ"
  k3 = "MQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQZ"
  k4 = "MQQQQZMQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQZMQQQQQZ"
  k5 = "MQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQZ"
  k6 = "MQQQQQQQQQZMQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQZ"
  k7 = "MQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQZ"
  k8 = "MQQQQQQQQZMQQQQQQQZMQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQZMQQQQQQQZ"
  k9 = "MQQQQQQQQZMQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQZ"
  ListPathAllKeywords = Array(a, b, c, d, e, f, g, h, i, j, k, l, m, n, o, p, q, r, s, t, u, v, w, X, Y, z, k0, k1, k2, k3, k4, k5, k6, k7, k8, k9)
End Function

Function listAllPathInSVG(ByVal svg$)
  Dim s$, p, i, z, k%, re, ms, m, m0!, m1!, m2!, v$, v0$, v1$, v2$
  Set re = glbRegex
  v = "(\d+(?:\.\d+)?)"
  're.Pattern = "" & v & "?([mavhlcqz])" & v & ""
  re.Pattern = "M.+?Z"
  p = Split(svg, " d=\""")
  ReDim z(1 To UBound(p) - 2)
  For i = 1 To UBound(p)
    s = Split(p(i), "\""")(0)
    Set ms = re.Execute(s): 'm0 = 0: m1 = 0
    If ms.count > 1 Then
'      Select Case ms.count
'      Case 2: s = ms(1)
'      Case 4: s = ms(2)
'      Case 6: s = ms(3)
'      End Select
      k = k + 1: z(k) = s ' ms(0)
    End If
'    For Each m In ms
'      v0 = m.submatches(0): v1 = m.submatches(1): v2 = m.submatches(2)
'      Select Case v1
'      Case "M", "m": m0 = 0 ' ÐiêÒm khõÒi hành
'      Case "A", "a": m0 = v2 ' VoÌng cung
'      Case "V", "v": m0 = v2 ' ÐýõÌng doòc
'      Case "H", "h": m0 = v2 ' ÐýõÌng ngang
'      Case "L", "l": m0 = v2 ' Phân ðoaòn
'      Case "C", "c", "Q", "q": If IsNumeric(v2) Then m0 = v2: ' ÐýõÌng cong
'      Case "Z", "z": m0 = v0: ' Ðóng
'      End Select
'n:
'    Next
  Next
  listAllPathInSVG = z
End Function


