<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>

<%
Dim maColors(139)
maColors(0) = "aliceblue^(#F0F8FF)^black"
maColors(1) = "antiquewhite^(#FAEBD7)^black"
maColors(2) = "aqua^(#00FFFF)^black"
maColors(3) = "aquamarine^(#7FFFD4)^black"
maColors(4) = "azure^(#F0FFFF)^black"
maColors(5) = "beige^(#F5F5DC)^black"
maColors(6) = "bisque^(#FFE4C4)^black"
maColors(7) = "black^(#000000)^white"
maColors(8) = "blanchedalmond^(#FFEBCD)^black"
maColors(9) = "blue^(#0000FF)^black"
maColors(10) = "blueviolet^(#8A2BE2)^black"
maColors(11) = "brown^(#A52A2A)^black"
maColors(12) = "burlywood^(#DEB887)^black"
maColors(13) = "cadetblue^(#5F9EA0)^black"
maColors(14) = "chartreuse^(#7FFF00)^black"
maColors(15) = "chocolate^(#D2691E)^black"
maColors(16) = "coral^(#FF7F50)^black"
maColors(17) = "cornflowerblue^(#6495ED)^black"
maColors(18) = "cornsilk^(#FFF8DC)^black"
maColors(19) = "crimson^(#DC143C)^black"
maColors(20) = "cyan^(#00FFFF)^black"
maColors(21) = "darkblue^(#00008B)^white"
maColors(22) = "darkcyan^(#008B8B)^black"
maColors(23) = "darkgoldenrod^(#B8860B)^black"
maColors(24) = "darkgray^(#A9A9A9)^black"
maColors(25) = "darkgreen^(#006400)^white"
maColors(26) = "darkkhaki^(#BDB76B)^black"
maColors(27) = "darkmagenta^(#8B008B)^white"
maColors(28) = "darkolivegreen^(#556B2F)^white"
maColors(29) = "darkorange^(#FF8C00)^black"
maColors(30) = "darkorchid^(#9932CC)^black"
maColors(31) = "darkred^(#8B0000)^white"
maColors(32) = "darksalmon^(#E9967A)^black"
maColors(33) = "darkseagreen^(#8FBC8B)^black"
maColors(34) = "darkslateblue^(#483D8B)^white"
maColors(35) = "darkslategray^(#2F4F4F)^white"
maColors(36) = "darkturquoise^(#00CED1)^black"
maColors(37) = "darkviolet^(#9400D3)^black"
maColors(38) = "deeppink^(#FF1493)^black"
maColors(39) = "deepskyblue^(#00BFFF)^black"
maColors(40) = "dimgray^(#696969)^white"
maColors(41) = "dodgerblue^(#1E90FF)^black"
maColors(42) = "firebrick^(#B22222)^black"
maColors(43) = "floralwhite^(#FFFAF0)^black"
maColors(44) = "forestgreen^(#228B22)^black"
maColors(45) = "fuchsia^(#FF00FF)^black"
maColors(46) = "gainsboro^(#DCDCDC)^black"
maColors(47) = "ghostwhite^(#F8F8FF)^black"
maColors(48) = "gold^(#FFD700)^black"
maColors(49) = "goldenrod^(#DAA520)^black"
maColors(50) = "gray^(#808080)^black"
maColors(51) = "green^(#008000)^white"
maColors(52) = "greenyellow^(#ADFF2F)^black"
maColors(53) = "honeydew^(#F0FFF0)^black"
maColors(54) = "hotpink^(#FF69B4)^black"
maColors(55) = "indianred^(#CD5C5C)^black"
maColors(56) = "indigo^(#4B0082)^white"
maColors(57) = "ivory^(#FFFFF0)^black"
maColors(58) = "khaki^(#F0E68C)^black"
maColors(59) = "lavender^(#E6E6FA)^black"
maColors(60) = "lavenderblush^(#FFF0F5)^black"
maColors(61) = "lawngreen^(#7CFC00)^black"
maColors(62) = "lemonchiffon^(#FFFACD)^black"
maColors(63) = "lightblue^(#ADD8E6)^black"
maColors(64) = "lightcoral^(#F08080)^black"
maColors(65) = "lightcyan^(#E0FFFF)^black"
maColors(66) = "lightgoldenrodyellow^(#FAFAD2)^black"
maColors(67) = "lightgreen^(#90EE90)^black"
maColors(68) = "lightgrey^(#D3D3D3)^black"
maColors(69) = "lightpink^(#FFB6C1)^black"
maColors(70) = "lightsalmon^(#FFA07A)^black"
maColors(71) = "lightseagreen^(#20B2AA)^black"
maColors(72) = "lightskyblue^(#87CEFA)^black"
maColors(73) = "lightslategray^(#778899)^black"
maColors(74) = "lightsteelblue^(#B0C4DE)^black"
maColors(75) = "lightyellow^(#FFFFE0)^black"
maColors(76) = "lime^(#00FF00)^black"
maColors(77) = "limegreen^(#32CD32)^black"
maColors(78) = "linen^(#FAF0E6)^black"
maColors(79) = "magenta^(#FF00FF)^black"
maColors(80) = "maroon^(#800000)^white"
maColors(81) = "mediumaquamarine^(#66CDAA)^black"
maColors(82) = "mediumblue^(#0000CD)^black"
maColors(83) = "mediumorchid^(#BA55D3)^black"
maColors(84) = "mediumpurple^(#9370DB)^black"
maColors(85) = "mediumseagreen^(#3CB371)^black"
maColors(86) = "mediumslateblue^(#7B68EE)^black"
maColors(87) = "mediumspringgreen^(#00FA9A)^black"
maColors(88) = "mediumturquoise^(#48D1CC)^black"
maColors(89) = "mediumvioletred^(#C71585)^black"
maColors(90) = "midnightblue^(#191970)^white"
maColors(91) = "mintcream^(#F5FFFA)^black"
maColors(92) = "mistyrose^(#FFE4E1)^black"
maColors(93) = "moccasin^(#FFE4B5)^black"
maColors(94) = "navajowhite^(#FFDEAD)^black"
maColors(95) = "navy^(#000080)^white"
maColors(96) = "oldlace^(#FDF5E6)^black"
maColors(97) = "olive^(#808000)^black"
maColors(98) = "olivedrab^(#6B8E23)^black"
maColors(99) = "orange^(#FFA500)^black"
maColors(100) = "orangered^(#FF4500)^black"
maColors(101) = "orchid^(#DA70D6)^black"
maColors(102) = "palegoldenrod^(#EEE8AA)^black"
maColors(103) = "palegreen^(#98FB98)^black"
maColors(104) = "paleturquoise^(#AFEEEE)^black"
maColors(105) = "palevioletred^(#DB7093)^black"
maColors(106) = "papayawhip^(#FFEFD5)^black"
maColors(107) = "peachpuff^(#FFDAB9)^black"
maColors(108) = "peru^(#CD853F)^black"
maColors(109) = "pink^(#FFC0CB)^black"
maColors(110) = "plum^(#DDA0DD)^black"
maColors(111) = "powderblue^(#B0E0E6)^black"
maColors(112) = "purple^(#800080)^black"
maColors(113) = "red^(#FF0000)^black"
maColors(114) = "rosybrown^(#BC8F8F)^black"
maColors(115) = "royalblue^(#4169E1)^black"
maColors(116) = "saddlebrown^(#8B4513)^white"
maColors(117) = "salmon^(#FA8072)^black"
maColors(118) = "sandybrown^(#F4A460)^black"
maColors(119) = "seagreen^(#2E8B57)^black"
maColors(120) = "seashell^(#FFF5EE)^black"
maColors(121) = "sienna^(#A0522D)^black"
maColors(122) = "silver^(#C0C0C0)^black"
maColors(123) = "skyblue^(#87CEEB)^black"
maColors(124) = "slateblue^(#6A5ACD)^black"
maColors(125) = "slategray^(#708090)^black"
maColors(126) = "snow^(#FFFAFA)^black"
maColors(127) = "springgreen^(#00FF7F)^black"
maColors(128) = "steelblue^(#4682B4)^black"
maColors(129) = "tan^(#D2B48C)^black"
maColors(130) = "teal^(#008080)^black"
maColors(131) = "thistle^(#D8BFD8)^black"
maColors(132) = "tomato^(#FF6347)^black"
maColors(133) = "turquoise^(#40E0D0)^black"
maColors(134) = "violet^(#EE82EE)^black"
maColors(135) = "wheat^(#F5DEB3)^black"
maColors(136) = "white^(#FFFFFF)^black"
maColors(137) = "whitesmoke^(#F5F5F5)^black"
maColors(138) = "yellow^(#FFFF00)^black"
maColors(139) = "yellowgreen^(#9ACD32)^black"

%>
<HTML><HEAD>
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<TITLE>Select Color</TITLE>
    <STYLE id=ThisPageStyles type="text/css" rel="stylesheet">
    </STYLE>
</HEAD>

<SCRIPT LANGUAGE=vbscript>
Dim mstrColor

Sub window_onload
    txtFocus.focus
    mstrColor = ""
End Sub

Sub SelectColor(strColor)
    Call window.parent.SelectColor(strColor)
End Sub

Sub MarkColor(strColor)
    mstrColor = strColor
End Sub

Sub Document_onkeypress()
    Select Case window.event.keyCode
        Case 27
            Call window.parent.SelectColor("")
        Case 13
            If mstrColor <> "" Then Call window.parent.SelectColor(mstrColor)
    End Select
End Sub

</SCRIPT>

<BODY style="BACKGROUND-COLOR:#white; overflow:scroll" bottomMargin=0 leftMargin=0 topMargin=0 rightMargin=0>
    <INPUT type="text" ID=txtFocus NAME="txtFocus" style="width:20;top:-100;position:absolute">
    <TABLE id=tblColors>
        <TBODY id=tbdColors>
            <%
            Dim intI
            Dim intJ
            Dim strColor
            Dim strHex
            Dim strTextColor
            
            intJ = -1
            For intI = 0 To 139
                If intI Mod 4 = 0 Then
                    If intI > 0 Then
                        Response.Write "</TR>"
                    End If
                    Response.Write "<TR id=tbr" & intI/4 & " style=""width:250"">"
                End If
                intJ = InStr(maColors(intI),"^")
                strColor = Parse(maColors(intI),"^",1) ' Mid(maColors(intI),1,intJ-1)
                strHex = Parse(maColors(intI),"^",2) ' Mid(maColors(intI),intJ+1,Len(maColors(intI))-intJ)
                strTextColor = Parse(maColors(intI),"^",3)
                Response.Write "<TD id=" & strColor & " onclick=MarkColor(""" & strColor & """) ondblclick=SelectColor(""" & strColor & """) style=""background-color:" & strColor & ";color:" & strTextColor & ";width:45;font-size:8pt;font-family:arial;overflow:hidden;height:20"">" & strColor & "</TD>"
            Next
            %>
        </TBODY>
    </TABLE>
</BODY>
</HTML>

<!--#include file="IncSvrFunctions.asp"-->