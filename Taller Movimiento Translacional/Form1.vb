Imports System.IO

Public Class Form1
    Dim k1, k2, k3, k4, b1, b2, b3 As Double
    Dim m1, m2 As Double
    Dim pInfo As New ProcessStartInfo
    Dim p As Process
    Dim vectorx1(10) As Double
    Dim vectorx2(10) As Double
    Dim vectort(10) As Double
    Dim posxm1, posxm2 As Integer
    Dim posxk2, posxk3, posxk4, posxb2, posxb3 As Integer
    Dim posxllanta1, posxllanta2, posxllanta3, posxllanta4, posxcarbon, posxcarbon2 As Integer
    Dim can_elementos, ganancia As Integer
    Dim tami1, tami5, tami8, tami9 As Integer
    Dim tipoGrafica, a1, a2, a3, a4, respuesta As String
    Dim aux As Integer = 0
    Dim ruta As String = "D:\PROGRAMAS_INSTALADOS\Octave\Octave-5.2.0\mingw64\bin\octave-cli-5.2.0.exe"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        pInfo.FileName = ruta
        pInfo.WindowStyle = ProcessWindowStyle.Minimized
        p = Process.Start(pInfo)
        posxm1 = PBmasa1.Location.X
        posxm2 = PBmasa2.Location.X
        posxllanta1 = PBllanta1.Location.X
        posxllanta2 = PBllanta2.Location.X
        posxllanta3 = PBllanta3.Location.X
        posxllanta4 = PBllanta4.Location.X
        posxcarbon = PBcarbon.Location.X
        posxcarbon2 = PBcarbon2.Location.X

        posxk2 = PBK2.Location.X
        posxk3 = PBK3.Location.X
        posxk4 = PBK4.Location.X
        posxb2 = PBb2.Location.X
        posxb3 = PBb3.Location.X
        tami1 = PBK1.Width
        tami5 = PBb1.Width
        tami8 = PBmasa2.Width
        tami9 = PBpared2.Width
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        k1 = TBK1.Text
        k2 = TBK2.Text
        k3 = TBK3.Text
        k4 = TBK4.Text
        b1 = TBb1.Text
        b2 = TBb2.Text
        b3 = TBb3.Text
        m1 = TBmasa1.Text
        m2 = TBmasa2.Text
        can_elementos = TBcantE.Text
        ganancia = TBgan.Text
        tipoGrafica = ComboBox1.Text
        Chart1.Titles.Add("FUNCION DE TRANSFERENCIA G1(S) - RESPUESTA AL " & tipoGrafica)
        Chart2.Titles.Add("FUNCION DE TRANSFERENCIA G2(S) - RESPUESTA AL " & tipoGrafica)
        If tipoGrafica = "PASO" Then
            respuesta = "step"
        Else
            respuesta = "impulse"
        End If

        a1 = m1 & "s^2 +" & k1 & " + " & b1 & "s" & " + " & b2 & "s +" & k2 & " + " & k3
        a2 = b2 & "s +" & k2 & " + " & k3
        a3 = m2 & "s^2 + " & b2 & "s +" & k2 & " + " & k3 & " + " & k4 & " + " & b3 & "s"
        a4 = b2 & "s + " & k2 & " + " & k3
        Labelnum1.Text = "(" & a3 & ")"
        Labelden1.Text = "(" & a1 & ")" & "(" & a3 & ")" & " - " & "(" & a4 & ")" & "(" & a2 & ")"
        Labelnum2.Text = "(" & a4 & ")"
        Labelden2.Text = "(" & a1 & ")" & "(" & a3 & ")" & " - " & "(" & a4 & ")" & "(" & a2 & ")"

        sendOctave("clc")
        sendOctave("clear")
        sendOctave("pkg load control")
        sendOctave("s=tf{(}'s'{)}")
        sendOctave("k1=" & k1 & ";")
        sendOctave("k2=" & k2 & ";")
        sendOctave("k3=" & k3 & ";")
        sendOctave("k4=" & k4 & ";")
        sendOctave("b1=" & b1 & ";")
        sendOctave("b2=" & b2 & ";")
        sendOctave("b3=" & b3 & ";")
        sendOctave("m1=" & m1 & ";")
        sendOctave("m2=" & m2 & ";")
        sendOctave("a1=m1*s*s{+}k1{+}b1*s{+}b2*s{+}k2{+}k3;")
        sendOctave("a2=b2*s{+}k2{+}k3;")
        sendOctave("a3=m2*s*s{+}b2*s{+}k2{+}k3{+}k4{+}b3*s;")
        sendOctave("a4=b2*s{+}k2{+}k3;")
        sendOctave("G1=a3/{(}a1*a3-a4*a2{)};")
        sendOctave("G2=G1*a4/a3;")

        sendOctave("[x1,t1]=" & respuesta & "{(}G1{)};")
        sendOctave("c=length{(}t1{)};")
        sendOctave("tiempo=t1{(}c{)}*1.1;")
        sendOctave("[x1,t1]=" & respuesta & "{(}G1{)},tiempo,tiempo/" & can_elementos & "{)};")
        sendOctave("dlmwrite{(}'" & Application.StartupPath & "\t1.txt',t1,'\n'{)};")
        sendOctave("dlmwrite{(}'" & Application.StartupPath & "\x1.txt',x1,'\n'{)};")

        sendOctave("[x2,t2]=" & respuesta & "{(}G2{)};")
        sendOctave("c=length{(}t2{)};")
        sendOctave("tiempo=t2{(}c{)}*1.1;")
        sendOctave("[x2,t2]=" & respuesta & "{(}G2{)},tiempo,tiempo/" & can_elementos & "{)};")
        sendOctave("dlmwrite{(}'" & Application.StartupPath & "\t2.txt',t2,'\n'{)};")
        sendOctave("dlmwrite{(}'" & Application.StartupPath & "\x2.txt',x2,'\n'{)};")
        sendOctave("exit")
        Cargar()
    End Sub

    Sub Cargar()
        ReDim vectorx1(can_elementos)
        ReDim vectorx2(can_elementos)
        ReDim vectort(can_elementos)
        Dim datosx1, datosx2, datost As StreamReader

        datosx1 = New StreamReader(Application.StartupPath & "\x1.txt")
        For j As Integer = 0 To can_elementos - 1
            vectorx1(j) = Val(datosx1.ReadLine) * ganancia
        Next
        datosx1.Close()

        datosx2 = New StreamReader(Application.StartupPath & "\x2.txt")
        For j As Integer = 0 To can_elementos - 1
            vectorx2(j) = Val(datosx2.ReadLine) * ganancia
        Next
        datosx2.Close()

        datost = New StreamReader(Application.StartupPath & "\t1.txt")
        For j As Integer = 0 To can_elementos - 1
            vectort(j) = Val(datost.ReadLine)
        Next
        datost.Close()

        Timer1.Enabled = True
    End Sub

    Sub sendOctave(cadena As String)
        AppActivate(ruta)
        SendKeys.SendWait(cadena & Chr(13))
        System.Threading.Thread.Sleep(100)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        PBmasa1.Location = New Point(posxm1 + vectorx1(aux), PBmasa1.Location.Y)
        PBmasa2.Location = New Point(posxm2 + vectorx2(aux), PBmasa2.Location.Y)
        PBllanta1.Location = New Point(posxllanta1 + vectorx1(aux), PBllanta1.Location.Y)
        PBllanta2.Location = New Point(posxllanta2 + vectorx1(aux), PBllanta2.Location.Y)
        PBllanta3.Location = New Point(posxllanta3 + vectorx2(aux), PBllanta3.Location.Y)
        PBllanta4.Location = New Point(posxllanta4 + vectorx2(aux), PBllanta4.Location.Y)
        PBcarbon.Location = New Point(posxcarbon + vectorx1(aux), PBcarbon.Location.Y)
        PBcarbon2.Location = New Point(posxcarbon2 + vectorx2(aux), PBcarbon2.Location.Y)

        PBb2.Location = New Point(posxb2 + vectorx1(aux), PBb2.Location.Y)
        PBK2.Location = New Point(posxk2 + vectorx1(aux), PBK2.Location.Y)
        PBK3.Location = New Point(posxk3 + vectorx1(aux), PBK3.Location.Y)
        PBK4.Location = New Point(posxk4 + vectorx2(aux), PBK4.Location.Y)
        PBb3.Location = New Point(posxb3 + vectorx2(aux), PBb3.Location.Y)

        PBK1.Width = tami1 + vectorx1(aux)
        PBb1.Width = tami5 + vectorx1(aux)
        PBb2.Width = PBmasa2.Location.X - PBmasa1.Location.X
        PBK2.Width = PBmasa2.Location.X - PBmasa1.Location.X
        PBK3.Width = PBmasa2.Location.X - PBmasa1.Location.X
        PBK4.Width = PBpared2.Location.X - PBmasa2.Location.X - tami8 + tami9
        PBb3.Width = PBpared2.Location.X - PBmasa2.Location.X - tami8 + tami9

        Chart1.Series(0).Points.AddXY(vectort(aux), vectorx1(aux))
        Chart2.Series(0).Points.AddXY(vectort(aux), vectorx2(aux))

        aux += 1
        If aux = can_elementos Then
            Timer1.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Application.Restart()
    End Sub
End Class
