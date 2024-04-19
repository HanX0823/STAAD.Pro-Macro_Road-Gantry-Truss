'#Reference {EDA9FA7F-EFC9-4264-9513-39CF6E72604D}#1.0#0#C:\Program Files\Bentley\Engineering\STAAD.Pro CONNECT Edition\STAAD\StaadPro.dll#OpenSTAADUI
'/*--------------------------------------------------------------------------------------+
'|  $Reference:
'|  https://docs.bentley.com/LiveContent/web/STAAD.Pro%20Help-v20/en/GUID-B726AC66-284B-4EEF-AFC0-AE695419D421.html
'|
'|  $Modified by Hanx0823. All rights reserved. $
'+--------------------------------------------------------------------------------------*/
Sub Main()
'DESCRIPTION:Create a 2D frame with supports

    Begin Dialog UserDialog 700,380,"STAAD.Pro Macro: Automatic Model Generation of LTA Road Gantry Truss with Traffic Signage" ' %GRID:5,5,1,1
        Text 20,20,190,15,"Width, W (m) =",.Text1
        Text 20,45,190,15,"Heigth, H (m) =",.Text2
        Text 20,70,190,15,"Span, L,max (m) =",.Text3
        Text 20,95,190,15,"Max L1 (mm) =",.Text4
        Text 20,120,190,15,"L2 (mm) =",.Text5
        Text 20,145,190,15,"L3 (mm) =",.Text6
        Text 20,170,190,15,"L4 (mm) =",.Text7
        Text 20,195,190,15,"L5 (mm) =",.Text8
        Text 20,220,190,15,"L6 (mm) =",.Text9
        Text 20,245,190,15,"L7 (mm) =",.Text10
        Text 20,270,190,15,"Signboard Distance (mm) =",.Text11
        Text 20,295,190,15,"Signboard Type:",.Text12
        Text 20,320,190,15,"Support Type:",.Text13
        TextBox 220,20,130,15,.W
        TextBox 220,45,130,15,.H
        TextBox 220,70,130,15,.Lmax
        TextBox 220,95,130,15,.L1
        TextBox 220,120,130,15,.L2
        TextBox 220,145,130,15,.L3
        TextBox 220,170,130,15,.L4
        TextBox 220,195,130,15,.L5
        TextBox 220,220,130,15,.L6
        TextBox 220,245,130,15,.L7
        TextBox 220,270,130,15,.SignDist
        OptionGroup .sign
            OptionButton 220,295,180,15,"Type 1 (3 x L6)",.OptionButton1
            OptionButton 370,295,180,15,"Type 2 (4 x L6)",.OptionButton2
            OptionButton 520,295,180,15,"Type 3 (7 x L6)",.OptionButton3
        OptionGroup .sprt
            OptionButton 220,320,90,15,"Fixed",.OptionButton4
            OptionButton 370,320,90,15,"Pinned",.OptionButton5
        OKButton 220,345,90,20
        CancelButton 370,345,90,20
    End Dialog
    Dim dlg As UserDialog

    Dim dlgResult As Integer
    Dim crdx As Double
    Dim crdy As Double
    Dim crdz As Double
    'Long is basically integer but with higher limits
    Dim n1 As Long
    Dim n2 As Long
    Dim i1 As Long
    Dim s1 As Long

    'No. of L1 Truss
    Dim nt As Integer
    Dim count As Integer

    'Beam numbers
    Dim bnum As Integer
    'Node numbers
    Dim nnum As Integer
    nnum = 1

    'Initialization
    dlg.W = "12.8"
    dlg.H = "5.24"
    dlg.Lmax = "14.24"
    dlg.L1 = "1400"
    dlg.L2 = "1800"
    dlg.L3 = "2500"
    dlg.L4 = "200"              'refer to LTA SDRE
    dlg.L5 = "200"              'refer to LTA SDRE
    dlg.L6 = "691.4286"
    dlg.L7 = "1400"             'refer to LTA SDRE
    dlg.SignDist = "2000"         'refer to LTA SDRE

    dlg.sprt = "1"         'PINNED because in Tedds CHS can't take moment, and in Hilti Profis CHS with moment sure FAIL kao kao.

    'Popup the dialog
    dlgResult = Dialog(dlg)
    Debug.Clear

    If dlgResult = -1 Then 'OK button pressed
        Debug.Print "OK button pressed"
        Debug.Print ""

        W = Abs( CDbl(dlg.W) )
        H = Abs( CDbl(dlg.H) )
        Lmax = Abs( CDbl(dlg.Lmax) )
        L1 = Abs( CDbl(dlg.L1) ) / 1000
        L2 = Abs( CDbl(dlg.L2) ) / 1000
        L2 = Abs( CDbl(dlg.L2) ) / 1000
        L3 = Abs( CDbl(dlg.L3) ) / 1000
        L4 = Abs( CDbl(dlg.L4) ) / 1000
        L5 = Abs( CDbl(dlg.L5) ) / 1000
        L6 = Abs( CDbl(dlg.L6) ) / 1000
        L7 = Abs( CDbl(dlg.L7) ) / 1000
        SignDist = Abs( CDbl(dlg.SignDist) ) / 1000
        sign = CStr(dlg.sign)
        sprt = CStr(dlg.sprt)

        Debug.Print "Info below are for debug:"
        Debug.Print "Wdith, W = ";W
        Debug.Print "Height, H = ";H
        Debug.Print "Span, Lmax = ";Lmax
        Debug.Print "Max L1 = ";L1
        Debug.Print "L2 = ";L2
        Debug.Print "L3 = ";L3
        Debug.Print "L4 = ";L4
        Debug.Print "L5 = ";L5
        Debug.Print "L6 = ";L6
        Debug.Print "L7 = ";L7
        Debug.Print "Signboard Type = ";sign
        Debug.Print "Support Type = ";sprt
        Debug.Print ""

        Set objOpenSTAAD = GetObject(,"StaadPro.OpenSTAAD")

        Dim geometry As OSGeometryUI
        Set geometry = objOpenSTAAD.Geometry

        '===============================================Part 1 - Truss Nodes & Members==================================================================
        'Origin Node 1=============================================================================================================================
        crdx = 0
        crdy = 0
        crdz = 0
        geometry.AddNode crdx, crdy, crdz
        Debug.Print "Node ";nnum;" added at x=";crdx;", y =";crdy;", z =";crdz
        nnum = nnum + 1

        'Node 2
        crdx = Lmax
        geometry.AddNode crdx, crdy, crdz
        Debug.Print "Node ";nnum;" added at x=";crdx;", y =";crdy;", z =";crdz
        nnum = nnum + 1

        'Node 3 - Stump
        crdx = 0
        crdy = 0.7
        geometry.AddNode crdx, crdy, crdz
        Debug.Print "Node ";nnum;" added at x=";crdx;", y =";crdy;", z =";crdz
        nnum = nnum + 1

        'Supports======================================================================================================================
        Dim support As OSSupportUI
        Set support = objOpenSTAAD.Support

        Debug.Print ""
        If sprt = "0" Then
            s1 = support.CreateSupportFixed()
            Debug.Print "sprt = 0"
        ElseIf sprt = "1" Then
            s1 = support.CreateSupportPinned()
            Debug.Print "sprt = 1"
        Else
            Debug.Print "sprt = error"
            MsgBox("Select Proper Support Type",vbOkOnly,"Error")
            Exit Sub
        End If
        Debug.Print "Support return value = ";s1
        Debug.Print ""
        For i1 = 1 To 2
            support.AssignSupportToNode i1,s1
        Next

        'Node 4 - Stump======================================================================================================================
        crdx = Lmax
        crdy = 0.7
        geometry.AddNode crdx, crdy, crdz
        Debug.Print "Node ";nnum;" added at x=";crdx;", y =";crdy;", z =";crdz
        nnum = nnum + 1

        'Node 5
        crdx = 0
        crdy = 5.7 - (L3-L2)
        geometry.AddNode crdx, crdy, crdz
        Debug.Print "Node ";nnum;" added at x=";crdx;", y =";crdy;", z =";crdz
        nnum = nnum + 1

        'Node 6
        crdx = Lmax
        crdy = 5.7 - (L3-L2)
        geometry.AddNode crdx, crdy, crdz
        Debug.Print "Node ";nnum;" added at x=";crdx;", y =";crdy;", z =";crdz
        nnum = nnum + 1

        'Node 7
        crdx = 0
        crdy = 5.7
        geometry.AddNode crdx, crdy, crdz
        Debug.Print "Node ";nnum;" added at x=";crdx;", y =";crdy;", z =";crdz
        nnum = nnum + 1

        'Node 8
        crdx = Lmax
        crdy = 5.7
        geometry.AddNode crdx, crdy, crdz
        Debug.Print "Node ";nnum;" added at x=";crdx;", y =";crdy;", z =";crdz
        nnum = nnum + 1

        'Node 9
        crdx = 0 + L2
        crdy = 5.7 + L2
        geometry.AddNode crdx, crdy, crdz
        Debug.Print "Node ";nnum;" added at x=";crdx;", y =";crdy;", z =";crdz
        nnum = nnum + 1

        'Node 10
        crdx = Lmax - L2
        crdy = 5.7 + L2
        geometry.AddNode crdx, crdy, crdz
        Debug.Print "Node ";nnum;" added at x=";crdx;", y =";crdy;", z =";crdz
        nnum = nnum + 1

        'Node 11
        crdx = 0 + L2
        crdy = 5.7
        geometry.AddNode crdx, crdy, crdz
        Debug.Print "Node ";nnum;" added at x=";crdx;", y =";crdy;", z =";crdz
        nnum = nnum + 1

        'Node 12
        crdx = Lmax - L2
        crdy = 5.7
        geometry.AddNode crdx, crdy, crdz
        Debug.Print "Node ";nnum;" added at x=";crdx;", y =";crdy;", z =";crdz
        nnum = nnum + 1

        'Truss Nodes from 13================================================================================================================================
        Dim ntreal As Double
        ntreal = (Lmax - (2 * L2)) / L1
        nt = Fix((Lmax - (2 * L2)) / L1) + 1                     'by Default, roundoff() To nearest Integer
        'nt = Fix((Lmax - (2 * L2)) / L1)
        ''in VBA, converting double or real numbers to integer numbers will use roundoff() function instead of trunc() function.
        ''in VBA, Fix() = Trunc() function.
        Debug.Print ""
        Debug.Print "No. of Truss (ntreal) = ";ntreal
        Debug.Print "No. of Truss (nt) = ";nt

        'If nt Mod 2 <> 0 Then                          'when nt is NOT an EVEN number
        '    nt = nt - 1
        'End If
        'Debug.Print "No. of Truss (nt) = ";nt
        'Debug.Print "" 

        L1 = (Lmax - (2 * L2)) / nt                     'calculating the real L1 value, using the values of given Max L1, L2 & Lmax.
        Debug.Print "real L1 value = ";L1
        Debug.Print ""

        'nnum = 13
        For j = 1 To 2                              'ROWs
            crdy = 5.7 + ((j - 1) * L2)
            'crdz = 1.0                             'move truss nodes towards z by 1.0m to separate them from other nodes during debugging
            For i = 1 To (nt - 1)                   'COLUMNs    No.of Nodes = No.of Truss - 1
                crdx = L2 + (i  * L1)
                geometry.AddNode crdx, crdy, crdz
                Debug.Print "Node ";nnum;" added at x=";crdx;", y =";crdy;", z =";crdz
                nnum = nnum + 1
            Next
        Next
        Debug.Print "========= Truss Nodes end here. ========="
        Debug.Print ""

        'Truss Beams 1 to 12=================================================================================================================
        n1 = 1
        n2 = 3
        bnum = 1
        For i = 1 To 8
            geometry.AddBeam n1, n2
            Debug.Print "add beam";bnum;" between ";n1;n2
            n1 = n1 + 1
            n2 = n2 + 1
            bnum = bnum + 1
        Next
        Debug.Print ""

        geometry.AddBeam 7, 11
        Debug.Print "add beam";bnum;" between ";7;11
        bnum = bnum + 1
        geometry.AddBeam 8, 12
        Debug.Print "add beam";bnum;" between ";8;12
        bnum = bnum + 1

        geometry.AddBeam 9, 11
        Debug.Print "add beam";bnum;" between ";9;11
        bnum = bnum + 1
        geometry.AddBeam 10, 12
        Debug.Print "add beam";bnum;" between ";10;12
        bnum = bnum + 1
        Debug.Print ""

        'Truss Beams 13 & 14=================================================================================================================
        geometry.AddBeam 11, 13
        Debug.Print "add beam";bnum;" between ";11;13
        bnum = bnum + 1

        Debug.Print "No. of Truss (nt) = ";nt
        n1 = 12
        n2 = 12 + (nt - 1)
        geometry.AddBeam n1, n2
        Debug.Print "add beam";bnum;" between ";n1;n2
        bnum = bnum + 1
        Debug.Print ""

        'Truss Beams 15 & 16=================================================================================================================
        n1 = 9
        n2 = 12 + nt
        geometry.AddBeam n1, n2
        Debug.Print "add beam";bnum;" between ";n1;n2
        bnum = bnum + 1

        n1 = 10
        n2 = 12 + 2*(nt -1)
        geometry.AddBeam n1, n2
        Debug.Print "add beam";bnum;" between ";n1;n2
        bnum = bnum + 1
        Debug.Print ""

        'Vertical Truss Beams from 17=================================================================================================================
        n1 = 13
        n2 = 13 + (nt - 1)
        For i = 1 To (nt - 1)
            geometry.AddBeam n1, n2
            Debug.Print "add beam";bnum;" between ";n1;n2
            n1 = n1 + 1
            n2 = n2 + 1
            bnum = bnum + 1
        Next
        Debug.Print ""

        'Bottom Truss Beams from 24=================================================================================================================
        n1 = 13
        n2 = n1 + 1
        For i = 1 To (nt - 2)
            geometry.AddBeam n1, n2
            Debug.Print "add beam";bnum;" between ";n1;n2
            n1 = n1 + 1
            n2 = n2 + 1
            bnum = bnum + 1
        Next
        Debug.Print ""

        'Top Truss Beams from 30=================================================================================================================
        n1 = 13 + (nt -1)
        n2 = n1 + 1
        For i = 1 To (nt - 2)
            geometry.AddBeam n1, n2
            Debug.Print "add beam";bnum;" between ";n1;n2
            n1 = n1 + 1
            n2 = n2 + 1
            bnum = bnum + 1
        Next
        Debug.Print ""
        Debug.Print "=========Truss beam end here. ========="
        Debug.Print ""

        '===============================================Part 1.5 - Bracing Members=====================================================================
        'No. of Bracing Members = nb
        Debug.Print "No. of Truss (nt) = ";nt
        Dim nberal As Double
        Dim nb As Integer
        nbreal = nt / 2
        nb = nt / 2                     'by default, this is using roundoff() to the nearest integer.
        Debug.Print "No.of Bracing Members (nbreal) =";nbreal
        Debug.Print "No.of Bracing Members (nb) =";nb
        Debug.Print ""

        'Truss Bracing Members from 36=================================================================================================================
        n1 = 9
        n2 = 13
        geometry.AddBeam n1, n2
        Debug.Print "add Bracing beam";bnum;" between ";n1;n2
        bnum = bnum + 1

        n1 = 13 + nt - 1
        n2 = 13 + 1
        For i = 1 To (nb - 1)
            geometry.AddBeam n1, n2
            Debug.Print "add Bracing beam";bnum;" between ";n1;n2
            n1 = n1 + 1
            n2 = n2 + 1
            bnum = bnum + 1
        Next
        Debug.Print ""

        'Truss Bracing Members from 40=================================================================================================================
        n1 = 13 + nt - ((nt / 2) + 1)
        n2 = 13 + nt + ((nt / 2) - 1)
        For i = 1 To (nb - 1)
            geometry.AddBeam n1, n2
            Debug.Print "add Bracing beam";bnum;" between ";n1;n2
            n1 = n1 + 1
            n2 = n2 + 1
            bnum = bnum + 1
        Next

        n1 = 13 + nt - 2
        n2 = 10
        geometry.AddBeam n1, n2
        Debug.Print "add Bracing beam";bnum;" between ";n1;n2
        bnum = bnum + 1
        Debug.Print "========= Bracing beam end here. ========="
        Debug.Print ""

        '===============================================Part 2 - Signboard Nodes & Members=======================================================
        'Signboard Nodes from 40=================================================================================================================
        '(2023-02-10): need to add truss nodes & members, signboard nodes & members, only then connection nodes (so split members will be created automatically),
        '(2023-02-10): followed by adding connection members (Member E between truss and signboard)
        crdx = L2           '1st Node for 1st row of Signboard Nodes
        crdy = 5.7 + 0.2    '1st Node for 1st row of Signboard Nodes
        crdz = SignDist      'The Signboard is 75mm away from the Truss (based on LTA/SDRE14/10/SUP10)
        'crdz = 0.075        'The Signboard is 75mm away from the Truss (based on LTA/SDRE14/10/SUP10)
        If sign = "0" Then
            nL6 = 3
        ElseIf sign = "1" Then
            nL6 = 4
        ElseIf sign = "2" Then
            nL6 = 7
        End If

        Dim CSBn As Integer                 'CSBn = 1st node of Connection (Signboard) Bop nodes
        For j = 1 To (nL6 + 1)                  'Horizontal beams of Signboard Nodes
            For i = 1 To (nt + 1)               'Vertical beams of Signboard Nodes
                If ((i = 1) And (j =1)) Then
                    CSBn = nnum
                End If
                geometry.AddNode crdx, crdy, crdz
                Debug.Print "Signboard Node ";nnum;" added at x=";crdx;", y =";crdy;", z =";crdz
                nnum = nnum + 1
                crdx = crdx + L1
                If (i > 1) Then
                    n1 = nnum - 2
                    n2 = nnum - 1
                    geometry.AddBeam n1, n2
                    Debug.Print "- add Signboard Horizontal beam";bnum;" between ";n1;n2
                    bnum = bnum + 1
                End If
                If (j > 1) Then
                    n1 = nnum - (nt + 1) - 1
                    n2 = nnum - 1
                    geometry.AddBeam n1, n2
                    Debug.Print "- add Signboard Vertical beam";bnum;" between ";n1;n2
                    bnum = bnum + 1
                End If
            Next
            crdx = L2
            crdy = crdy + L6
        Next
        Debug.Print "========= Signboard Nodes, Horizontal & Vertical beam end here. ========="
        Debug.Print ""

        '===============================================Part 3 - Connection Nodes & Members==================================================================
        'Connection (Truss) Nodes -(Preparation)================================================================================================================
        'No.of Signboard Frame (Member d)(nd) (Max Member d Span is set as 1400mm based on LTA SDRE drawings: LTA/SDRE14/10/SUP10)
        'calculating the No.of Signboard Frame (Member d)(nd), using the values of given Width of Signbaord (W) and Max Member d Span = 1400mm.
        Dim ndreal As Double
        Dim nd As Integer
        ndreal = W / 1.4
        nd = Fix(W / 1.4) + 1                                 'by Default, roundoff() to the nearest Integer
        'in VBA, Fix() = Trunc() function.
        Debug.Print "No. of Signboard (Member d) (ndreal) = ";ndreal
        Debug.Print "No. of Signboard (Member d) (nd) = ";nd

        'calculating the Length of Signboard (Member d)(Ld), using the values of given Width of Signbaord (W) and No.of Signboard Frame (Member d)(nd).
        Dim Ld As Double
        Ld = W / nd                                     'calculating the real L1 value, using the values of given Max L1, L2 & Lmax.
        Debug.Print "Length of Signboard (Member d)(Ld) = ";Ld
        Debug.Print ""

        'No.of Signboard Nodes = No.of Signboard Frame (Member d)(nd) + 1
        Dim ndnodes As Integer
        ndnodes = nd + 1

        'Define No.of L6 (nL6) based on Signboard Type
        If sign = "0" Then
            nL6 = 3
        ElseIf sign = "1" Then
            nL6 = 4
        ElseIf sign = "2" Then
            nL6 = 7
        End If
        Debug.Print "Support Type = ";sprt
        Debug.Print "No. of L6 (nLd) = ";nL6

        'Calulate Length of L6 using Height of Signboard (H) and No.of L6 (nL6)
        L6 = (H - L4 - L5) / nL6
        Debug.Print "Length of L6 (L6) = ";L6

        'Connection (Truss) Bottom Nodes from 63=================================================================================================================
        crdz = 0                 '1st Node for 1st row of Connection Nodes
        crdy = 5.7 + 0.2    '1st Node for 1st row of Connection Nodes
        crdx = L2               '1st Node for 1st row of Connection Nodes
        Dim CTBn As Integer             'CTBn = 1st node of Connection (Truss) Bottom nodes
        For i = 1 To (nt + 1)
            If i = 1 Then
                CTBn = nnum
            End If
            geometry.AddNode crdx, crdy, crdz
            Debug.Print "Connection (Truss) Bottom Node ";nnum;" added at x=";crdx;", y =";crdy;", z =";crdz
            nnum = nnum + 1
            crdx = crdx + L1
        Next
        'Connection (Truss) Top Nodes from 72================================================
        crdy = 5.7 + L2 - 0.2    '1st Node for 2nd row of Connection Nodes
        crdx = L2                       '1st Node for 2nd row of Connection Nodes
        Dim CTTn As Integer             'CTTn = 1st node of Connection (Truss) Top nodes
        For i = 1 To (nt + 1)
            If i = 1 Then
                CTTn = nnum
            End If
            geometry.AddNode crdx, crdy, crdz
            Debug.Print "Connection (Truss) Top Node ";nnum;" added at x=";crdx;", y =";crdy;", z =";crdz
            nnum = nnum + 1
            crdx = crdx + L1
        Next
        Debug.Print ""

        'Connection (Signboard) Bottom Nodes  [OMITTED because always clash with Signboard Bottom Nodes]=====================================================================================
        'crdz = SignDist         'The Signboard is 75mm away from the Truss (based on LTA/SDRE14/10/SUP10)
        ''crdz = 0.075         'The Signboard is 75mm away from the Truss (based on LTA/SDRE14/10/SUP10)
        'crdy = 5.7 + 0.2    '1st Node for 1st row of Connection Nodes
        'crdx = L2           '1st Node for 1st row of Connection Nodes
        'For i = 1 To (nt + 1)
            'geometry.AddNode crdx, crdy, crdz
            'Debug.Print "Connection (Signboard) Bottom Nodes ";nnum;" added at x=";crdx;", y =";crdy;", z =";crdz
            'nnum = nnum + 1
            'crdx = crdx + L1
        'Next
        'Connection (Signboard) Top Nodes from 90================================================
        crdz = SignDist        'The Signboard is 75mm away from the Truss (based on LTA/SDRE14/10/SUP10)
        crdy = 5.7 + L2 - 0.2    '1st Node for 2nd row of Connection Nodes
        crdx = L2           '1st Node for 2nd row of Connection Nodes
        Dim CSTn As Integer             'CTTn = 1st node of Connection (Signboard) Top nodes
        For i = 1 To (nt + 1)
            If i = 1 Then
                CSTn = nnum
            End If
            geometry.AddNode crdx, crdy, crdz
            Debug.Print "Connection (Signboard) Top Nodes ";nnum;" added at x=";crdx;", y =";crdy;", z =";crdz
            nnum = nnum + 1
            crdx = crdx + L1
        Next
        Debug.Print ""

        '===============================================Part 3.5 - Connection Members==========================================================================
        '(Bottom) Connection Members (Member E) from 130===========================================================================================================
        n1 = CSBn            'CTTn = 1st node of Connection (Signboard) Bottom nodes
        n2 = CTBn            'CTTn = 1st node of Connection (Truss) Bottom nodes
        Debug.Print "========= (Bottom) Connection Members (Member E) start here. ========="
        For i = 1 To (nt + 1)
            geometry.AddBeam n1, n2
            Debug.Print i;". add (Bottom) Connection Members (Member E) as beam";bnum;" between Node ";n1;" and Node";n2
            n1 = n1 + 1
            n2 = n2 + 1
            bnum = bnum + 1
        Next
        Debug.Print "========= (Bottom) Connection Members (Member E) end here. ========="
        Debug.Print ""

        '(Top) Connection Members (Member E) from 140===========================================================================================================
        n1 = CTTn            'CTTn = 1st node of Connection (Truss) Top nodes
        n2 = CSTn            'CTTn = 1st node of Connection (Signboard) Top nodes
        Debug.Print "========= (Top) Connection Members (Member E) start here. ========="
        For i = 1 To (nt + 1)
            geometry.AddBeam n1, n2
            Debug.Print i;". add (Top) Connection Members (Member E) as beam";bnum;" between Node ";n1;" and Node";n2
            n1 = n1 + 1
            n2 = n2 + 1
            bnum = bnum + 1
        Next
        Debug.Print "========= (Top) Connection Members (Member E) end here. ========="
        Debug.Print ""

        '===============================================Geometry Tab ended===================================================================================
        

    ElseIf dlgResult = 0 Then 'Cancel button pressed
        Debug.Print "Cancel button pressed"
    End If

End Sub
