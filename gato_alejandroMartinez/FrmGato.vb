Public Class frmGato
    Dim j1, j2, ganador, jugador As String
    Dim turno As Integer
    Dim v1, v2, v3, v4, v5, v6, v7, v8, v9 As String
    Dim l7, l8 As Boolean
    Dim ganoCpu As Boolean = False
    Dim defenderCpu As Boolean = False

    Private Sub InstruccionesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InstruccionesToolStripMenuItem.Click
        MsgBox("1.- Ingresa al menú Archivo de la parte superior izquierda del formulario." & Chr(13) & Chr(13) &
       "2.- Selecciona un nuevo juego." & Chr(13) & Chr(13) &
       "3.- Escoge si juegas 1 vs 1 o juegas 1 vs CPU." & Chr(13) & Chr(13) &
       "Si escoges jugar 1 vs 1." & Chr(13) & Chr(13) &
       "  - Ingresa el nombre del jugador 1." & Chr(13) & Chr(13) &
       "  - Ingresa el nombre del jugador 2." & Chr(13) & Chr(13) &
       "Si escoges jugar 1 vs CPU." & Chr(13) & Chr(13) &
       "  - Ingresa tu nombre." & Chr(13) & Chr(13) &
       "Disfruta tu partida.", vbInformation, "Instrucciones.")
    End Sub

    Private Sub AutorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AutorToolStripMenuItem.Click
        MsgBox("Creditos del Autor." & Chr(13) & Chr(13) &
            "Alumno: Alejandro Martínez Rivera" & Chr(13) & Chr(13) &
            "Materia: Programación lógica y funcional" & Chr(13) & Chr(13) &
            "Software Visual Basic 2019" & Chr(13) & Chr(13) &
            "Carrera Ing. en Sistems" & Chr(13) & Chr(13) &
            "Octavo Semestre" & Chr(13) & Chr(13), vbInformation, "Creditos")
    End Sub

    Private Sub Vs1ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Vs1ToolStripMenuItem.Click
        'Inicia un nuevo juego

        'Hacemos invisibles las barras de ganar
        linea1.Visible = False
        linea2.Visible = False
        linea3.Visible = False
        linea4.Visible = False
        linea5.Visible = False
        linea6.Visible = False

        l7 = False
        lineaDiagonal7()

        l8 = False
        lineaDiagonal8()

        'reiniciamos los objetos y variables
        c1.Visible = False
        c2.Visible = False
        c3.Visible = False
        c4.Visible = False
        c5.Visible = False
        c6.Visible = False
        c7.Visible = False
        c8.Visible = False
        c9.Visible = False

        x1.Visible = False
        x2.Visible = False
        x3.Visible = False
        x4.Visible = False
        x5.Visible = False
        x6.Visible = False
        x7.Visible = False
        x8.Visible = False
        x9.Visible = False

        btn1.Visible = True
        btn2.Visible = True
        btn3.Visible = True
        btn4.Visible = True
        btn5.Visible = True
        btn6.Visible = True
        btn7.Visible = True
        btn8.Visible = True
        btn9.Visible = True

        btn1.Enabled = True
        btn2.Enabled = True
        btn3.Enabled = True
        btn4.Enabled = True
        btn5.Enabled = True
        btn6.Enabled = True
        btn7.Enabled = True
        btn8.Enabled = True
        btn9.Enabled = True

        v1 = ""
        v2 = ""
        v3 = ""
        v4 = ""
        v5 = ""
        v6 = ""
        v7 = ""
        v8 = ""
        v9 = ""

        turno = 1

        ganador = ""

        Do
            j1 = InputBox("Ingrese el nombre del primer jugador.", "Jugador 1")
            If Len(j1) < 3 Then
                MsgBox("El nombre del jugador no puede estar vacío y minimo 3 letras", vbCritical, "Invalid input.")
            Else
                Exit Do
            End If
        Loop While (True)

        Do
            j2 = InputBox("Ingrese el nombre del primer jugador.", "Jugador 1")
            If Len(j2) < 3 Then
                MsgBox("El nombre del jugador no puede estar vacío y minimo 3 letras", vbCritical, "Invalid input.")
            Else
                Exit Do
            End If
        Loop While (True)
        name1.Text = j1
        name2.Text = j2
        lblPlayer.Text = j1
    End Sub

    Private Sub MúsicaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MúsicaToolStripMenuItem.Click
        'Cancion activada
        If MúsicaToolStripMenuItem.Checked = False Then
            My.Computer.Audio.Play(My.Resources.ticTacToe_sound, AudioPlayMode.BackgroundLoop)
            MúsicaToolStripMenuItem.Checked = True
        Else
            My.Computer.Audio.Stop()
            MúsicaToolStripMenuItem.Checked = False
        End If
    End Sub



    Private Sub SalirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem.Click
        If (MsgBox("Deseas salir del juego?", vbYesNo, "Salir") = vbYes) Then
            End
        End If
    End Sub


    Sub checarGanador()
        'Este procedimiento checa si alguien gano o quedo empatado
        'Validamos las opciones de ganar del primer jugador osea las X
        If v1 = "X" And v2 = "X" And v3 = "X" Then
            ganador = j1
            linea1.Visible = True
            MsgBox("Gano " & j1, vbInformation, "Felicidades.")
            deshabilitarBotones()
            Exit Sub
        End If

        If v4 = "X" And v5 = "X" And v6 = "X" Then
            ganador = j1
            linea2.Visible = True
            MsgBox("Gano " & j1, vbInformation, "Felicidades.")
            deshabilitarBotones()
            Exit Sub
        End If

        If v7 = "X" And v8 = "X" And v9 = "X" Then
            ganador = j1
            linea3.Visible = True
            MsgBox("Gano " & j1, vbInformation, "Felicidades.")
            deshabilitarBotones()
            Exit Sub
        End If

        If v1 = "X" And v4 = "X" And v7 = "X" Then
            ganador = j1
            linea6.Visible = True
            MsgBox("Gano " & j1, vbInformation, "Felicidades.")
            deshabilitarBotones()
            Exit Sub
        End If

        If v2 = "X" And v5 = "X" And v8 = "X" Then
            ganador = j1
            linea5.Visible = True
            MsgBox("Gano " & j1, vbInformation, "Felicidades.")
            deshabilitarBotones()
            Exit Sub
        End If

        If v3 = "X" And v6 = "X" And v9 = "X" Then
            ganador = j1
            linea4.Visible = True
            MsgBox("Gano " & j1, vbInformation, "Felicidades.")
            deshabilitarBotones()
            Exit Sub
        End If

        If v1 = "X" And v5 = "X" And v9 = "X" Then
            ganador = j1
            l7 = True
            lineaDiagonal7()
            MsgBox("Gano " & j1, vbInformation, "Felicidades.")
            deshabilitarBotones()
            Exit Sub
        End If

        If v3 = "X" And v5 = "X" And v7 = "X" Then
            ganador = j1
            l8 = True
            lineaDiagonal8()
            MsgBox("Gano " & j1, vbInformation, "Felicidades.")
            deshabilitarBotones()
            Exit Sub
        End If


        'Validamos las opciones de ganar del segundo jugador osea los 0
        If v1 = "0" And v2 = "0" And v3 = "0" Then
            ganador = j2
            linea1.Visible = True
            MsgBox("Gano " & j2, vbInformation, "Felicidades.")
            deshabilitarBotones()
            Exit Sub
        End If

        If v4 = "0" And v5 = "0" And v6 = "0" Then
            ganador = j2
            linea2.Visible = True
            MsgBox("Gano " & j2, vbInformation, "Felicidades.")
            deshabilitarBotones()
            Exit Sub
        End If

        If v7 = "0" And v8 = "0" And v9 = "0" Then
            ganador = j2
            linea3.Visible = True
            MsgBox("Gano " & j2, vbInformation, "Felicidades.")
            deshabilitarBotones()
            Exit Sub
        End If

        If v1 = "0" And v4 = "0" And v7 = "0" Then
            ganador = j2
            linea6.Visible = True
            MsgBox("Gano " & j2, vbInformation, "Felicidades.")
            deshabilitarBotones()
            Exit Sub
        End If

        If v2 = "0" And v5 = "0" And v8 = "0" Then
            ganador = j2
            linea5.Visible = True
            MsgBox("Gano " & j2, vbInformation, "Felicidades.")
            deshabilitarBotones()
            Exit Sub
        End If

        If v3 = "0" And v6 = "0" And v9 = "0" Then
            ganador = j2
            linea4.Visible = True
            MsgBox("Gano " & j2, vbInformation, "Felicidades.")
            deshabilitarBotones()
            Exit Sub
        End If

        If v1 = "0" And v5 = "0" And v9 = "0" Then
            ganador = j2
            l7 = True
            lineaDiagonal7()
            MsgBox("Gano " & j2, vbInformation, "Felicidades.")
            deshabilitarBotones()
            Exit Sub
        End If

        If v3 = "0" And v5 = "0" And v7 = "0" Then
            ganador = j2
            l8 = True
            lineaDiagonal8()
            MsgBox("Gano " & j2, vbInformation, "Felicidades.")
            deshabilitarBotones()
            Exit Sub
        End If

        'Ahora checamos si quedo empatado entonces ganaría el gato
        If ganador = "" And v1 <> "" And v2 <> "" And v3 <> "" And v4 <> "" And v5 <> "" _
            And v6 <> "" And v7 <> "" And v8 <> "" And v9 <> "" Then
            ganador = "El gato"
            MsgBox("Gano " & ganador, vbInformation, "Juego empatado")
            Exit Sub
        End If
    End Sub

    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
        btn1.Visible = False
        If turno = 1 Then
            x1.Visible = True
            v1 = "X"
        Else
            c1.Visible = True
            v1 = "0"
        End If

        checarGanador()
        cambiarTurno()

    End Sub

    Private Sub btn2_Click(sender As Object, e As EventArgs) Handles btn2.Click
        btn2.Visible = False
        If turno = 1 Then
            x2.Visible = True
            v2 = "X"
        Else
            c2.Visible = True
            v2 = "0"
        End If

        checarGanador()
        cambiarTurno()
    End Sub

    Private Sub btn3_Click(sender As Object, e As EventArgs) Handles btn3.Click
        btn3.Visible = False
        If turno = 1 Then
            x3.Visible = True
            v3 = "X"
        Else
            c3.Visible = True
            v3 = "0"
        End If

        checarGanador()
        cambiarTurno()
    End Sub

    Private Sub btn4_Click(sender As Object, e As EventArgs) Handles btn4.Click
        btn4.Visible = False
        If turno = 1 Then
            x4.Visible = True
            v4 = "X"
        Else
            c4.Visible = True
            v4 = "0"
        End If

        checarGanador()
        cambiarTurno()
    End Sub

    Private Sub btn5_Click(sender As Object, e As EventArgs) Handles btn5.Click
        btn5.Visible = False
        If turno = 1 Then
            x5.Visible = True
            v5 = "X"
        Else
            c5.Visible = True
            v5 = "0"
        End If

        checarGanador()
        cambiarTurno()
    End Sub

    Private Sub btn6_Click(sender As Object, e As EventArgs) Handles btn6.Click
        btn6.Visible = False
        If turno = 1 Then
            x6.Visible = True
            v6 = "X"
        Else
            c6.Visible = True
            v6 = "0"
        End If

        checarGanador()
        cambiarTurno()
    End Sub
    Private Sub btn7_Click(sender As Object, e As EventArgs) Handles btn7.Click
        btn7.Visible = False
        If turno = 1 Then
            x7.Visible = True
            v7 = "X"
        Else
            c7.Visible = True
            v7 = "0"
        End If

        checarGanador()
        cambiarTurno()
    End Sub

    Private Sub btn8_Click(sender As Object, e As EventArgs) Handles btn8.Click
        btn8.Visible = False
        If turno = 1 Then
            x8.Visible = True
            v8 = "X"
        Else
            c8.Visible = True
            v8 = "0"
        End If

        checarGanador()
        cambiarTurno()
    End Sub

    Private Sub btn9_Click(sender As Object, e As EventArgs) Handles btn9.Click
        btn9.Visible = False
        If turno = 1 Then
            x9.Visible = True
            v9 = "X"
        Else
            c9.Visible = True
            v9 = "0"
        End If

        checarGanador()
        cambiarTurno()
    End Sub

    Sub deshabilitarBotones()
        btn1.Enabled = False
        btn2.Enabled = False
        btn3.Enabled = False
        btn4.Enabled = False
        btn5.Enabled = False
        btn6.Enabled = False
        btn7.Enabled = False
        btn8.Enabled = False
        btn9.Enabled = False
    End Sub

    Sub cambiarTurno()
        If ganador = "" Then
            If turno = 1 Then
                turno = 2
                lblPlayer.Text = j2
                If j2 = "CPU" Then
                    posibilidad_ganar()
                    If ganoCpu = True Then
                        checarGanador()
                        defenderCpu = True
                        Exit Sub
                    End If
                    defender()
                    If defenderCpu = False And ganoCpu = False Then
                        casillaAleatoria()
                    End If
                    defenderCpu = False
                End If
            Else
                turno = 1
                lblPlayer.Text = j1
                If j1 = "CPU" Then
                    posibilidad_ganar()
                    If ganoCpu = True Then
                        checarGanador()
                        defenderCpu = True
                        Exit Sub
                    End If
                    defender()
                    If defenderCpu = False And ganoCpu = False Then
                        casillaAleatoria()
                    End If
                    defenderCpu = False
                End If
            End If
        End If
    End Sub


    'Diagonales

    'Metodo para crear la diagonal 7

    Sub lineaDiagonal7()
        If l7 = True Then
            x1.SendToBack()
            x5.SendToBack()
            x9.SendToBack()

            c1.SendToBack()
            c5.SendToBack()
            c9.SendToBack()

            PictureBox1.SendToBack()
            PictureBox2.SendToBack()
            PictureBox3.SendToBack()
            PictureBox4.SendToBack()

            LineShape1.Visible = True
            LineShape1.BringToFront()

        ElseIf l7 = False Then
            LineShape1.Visible = False
            LineShape1.SendToBack()
            PictureBox1.BringToFront()
            PictureBox2.BringToFront()
            PictureBox3.BringToFront()
            PictureBox4.BringToFront()

        End If

    End Sub

    'Metodo para crear la diagonal 8
    Sub lineaDiagonal8()
        If l8 = True Then
            x3.SendToBack()
            x5.SendToBack()
            x7.SendToBack()

            c3.SendToBack()
            c5.SendToBack()
            c7.SendToBack()

            PictureBox1.SendToBack()
            PictureBox2.SendToBack()
            PictureBox3.SendToBack()
            PictureBox4.SendToBack()

            LineShape2.Visible = True
            LineShape2.BringToFront()
        ElseIf l8 = False Then
            LineShape2.Visible = False
            LineShape2.SendToBack()
            PictureBox1.BringToFront()
            PictureBox2.BringToFront()
            PictureBox3.BringToFront()
            PictureBox4.BringToFront()
        End If

    End Sub

    Private Sub VsCPUToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VsCPUToolStripMenuItem.Click
        Dim turnoAleatorio As Integer
        ganoCpu = False
        defenderCpu = False
        'Inicia un nuevo juego

        'Hacemos invisibles las barras de ganar
        linea1.Visible = False
        linea2.Visible = False
        linea3.Visible = False
        linea4.Visible = False
        linea5.Visible = False
        linea6.Visible = False

        l7 = False
        lineaDiagonal7()

        l8 = False
        lineaDiagonal8()

        'reiniciamos los objetos y variables
        c1.Visible = False
        c2.Visible = False
        c3.Visible = False
        c4.Visible = False
        c5.Visible = False
        c6.Visible = False
        c7.Visible = False
        c8.Visible = False
        c9.Visible = False

        x1.Visible = False
        x2.Visible = False
        x3.Visible = False
        x4.Visible = False
        x5.Visible = False
        x6.Visible = False
        x7.Visible = False
        x8.Visible = False
        x9.Visible = False

        btn1.Visible = True
        btn2.Visible = True
        btn3.Visible = True
        btn4.Visible = True
        btn5.Visible = True
        btn6.Visible = True
        btn7.Visible = True
        btn8.Visible = True
        btn9.Visible = True

        btn1.Enabled = True
        btn2.Enabled = True
        btn3.Enabled = True
        btn4.Enabled = True
        btn5.Enabled = True
        btn6.Enabled = True
        btn7.Enabled = True
        btn8.Enabled = True
        btn9.Enabled = True

        v1 = ""
        v2 = ""
        v3 = ""
        v4 = ""
        v5 = ""
        v6 = ""
        v7 = ""
        v8 = ""
        v9 = ""

        turno = 1
        jugador = ""
        ganador = ""

        Do
            jugador = InputBox("Ingrese el nombre del jugador.", "Jugador")
            If Len(jugador) < 3 Then
                MsgBox("El nombre del jugador no puede estar vacío y minimo 3 letras", vbCritical, "Invalid input.")
            Else
                Randomize()
                turnoAleatorio = CInt(Rnd() * 2)
                If turnoAleatorio = 1 Then
                    MsgBox("Jugador 1: " & jugador & Chr(13) & Chr(13) &
                           "Jugador 2: CPU", vbInformation, "Eres el jugador 1")
                    j1 = jugador
                    j2 = "CPU"
                    name1.Text = jugador
                    name2.Text = "CPU"
                ElseIf turnoAleatorio = 2
                    MsgBox("Jugador 1: CPU" & Chr(13) & Chr(13) &
                            "Jugador 2: " & jugador, vbInformation, "Eres el jugador 2")
                    j1 = "CPU"
                    j2 = jugador
                    name1.Text = "CPU"
                    name2.Text = jugador
                End If
                lblPlayer.Text = j1
                If turnoAleatorio = 2 Then
                    casillaAleatoria()
                End If
                Exit Do
            End If
        Loop While (True)

    End Sub
    Private Sub casillaAleatoria()
        Dim casilla As Integer
        Randomize()
        casilla = CInt(Rnd() * 9)
        If casilla = 0 Then
            casillaAleatoria()
        End If

        If casilla = 1 And btn1.Visible = True Then
            btn1.Visible = False
            btn1.Enabled = False
            If turno = 1 Then
                x1.Visible = True
                v1 = "X"
                checarGanador()
                cambiarTurno()
            ElseIf turno = 2
                c1.Visible = True
                v1 = "0"
                checarGanador()
                cambiarTurno()
            End If
        ElseIf casilla = 1 And btn1.Visible = False
            casillaAleatoria()
        End If
        If casilla = 2 And btn2.Visible = True Then
            btn2.Visible = False
            btn2.Enabled = False
            If turno = 1 Then
                x2.Visible = True
                v2 = "X"
                checarGanador()
                cambiarTurno()
            ElseIf turno = 2
                c2.Visible = True
                v2 = "0"
                checarGanador()
                cambiarTurno()
            End If
        ElseIf casilla = 2 And btn2.Visible = False
            casillaAleatoria()
        End If
        If casilla = 3 And btn3.Visible = True Then
            btn3.Visible = False
            btn3.Enabled = False
            If turno = 1 Then
                x3.Visible = True
                v3 = "X"
                checarGanador()
                cambiarTurno()
            ElseIf turno = 2
                c3.Visible = True
                v3 = "0"
                checarGanador()
                cambiarTurno()
            End If
        ElseIf casilla = 3 And btn3.Visible = False
            casillaAleatoria()
        End If
        If casilla = 4 And btn4.Visible = True Then
            btn4.Visible = False
            btn4.Enabled = False
            If turno = 1 Then
                x4.Visible = True
                v4 = "X"
                checarGanador()
                cambiarTurno()
            ElseIf turno = 2
                c4.Visible = True
                v4 = "0"
                checarGanador()
                cambiarTurno()
            End If
        ElseIf casilla = 4 And btn4.Visible = False
            casillaAleatoria()
        End If
        If casilla = 5 And btn5.Visible = True Then
            btn5.Visible = False
            btn5.Enabled = False
            If turno = 1 Then
                x5.Visible = True
                v5 = "X"
                checarGanador()
                cambiarTurno()
            ElseIf turno = 2
                c5.Visible = True
                v5 = "0"
                checarGanador()
                cambiarTurno()
            End If
        ElseIf casilla = 5 And btn5.Visible = False
            casillaAleatoria()
        End If
        If casilla = 6 And btn6.Visible = True Then
            btn6.Visible = False
            btn6.Enabled = False
            If turno = 1 Then
                x6.Visible = True
                v6 = "X"
                checarGanador()
                cambiarTurno()
            ElseIf turno = 2
                c6.Visible = True
                v6 = "0"
                checarGanador()
                cambiarTurno()
            End If
        ElseIf casilla = 6 And btn6.Visible = False
            casillaAleatoria()
        End If
        If casilla = 7 And btn7.Visible = True Then
            btn7.Visible = False
            btn7.Enabled = False
            If turno = 1 Then
                x7.Visible = True
                v7 = "X"
                checarGanador()
                cambiarTurno()
            ElseIf turno = 2
                c7.Visible = True
                v7 = "0"
                checarGanador()
                cambiarTurno()
            End If
        ElseIf casilla = 7 And btn7.Visible = False
            casillaAleatoria()
        End If
        If casilla = 8 And btn8.Visible = True Then
            btn8.Visible = False
            btn8.Enabled = False
            If turno = 1 Then
                x8.Visible = True
                v8 = "X"
                checarGanador()
                cambiarTurno()
            ElseIf turno = 2
                c8.Visible = True
                v8 = "0"
                checarGanador()
                cambiarTurno()
            End If
        ElseIf casilla = 8 And btn8.Visible = False
            casillaAleatoria()
        End If
        If casilla = 9 And btn9.Visible = True Then
            btn9.Visible = False
            btn9.Enabled = False
            If turno = 1 Then
                x9.Visible = True
                v9 = "X"
                checarGanador()
                cambiarTurno()
            ElseIf turno = 2
                c9.Visible = True
                v9 = "0"
                checarGanador()
                cambiarTurno()
            End If
        ElseIf casilla = 9 And btn9.Visible = False
            casillaAleatoria()
        End If
    End Sub

    Private Sub posibilidad_ganar()
        If j1 = "CPU" Then

            If v2 = "X" And v3 = "X" And btn1.Visible = True Then
                btn1.Visible = False
                x1.Visible = True
                v1 = "X"
                ganoCpu = True
            ElseIf v5 = "X" And v9 = "X" And btn1.Visible = True
                btn1.Visible = False
                x1.Visible = True
                v1 = "X"
                ganoCpu = True
            ElseIf v4 = "X" And v7 = "X" And btn1.Visible = True
                btn1.Visible = False
                x1.Visible = True
                v1 = "X"
                ganoCpu = True
                If v1 = "X" And v3 = "X" And btn2.Visible = True Then
                    btn2.Visible = False
                    x2.Visible = True
                    v2 = "X"
                    ganoCpu = True
                ElseIf v5 = "X" And v8 = "X" And btn2.Visible = True
                    btn2.Visible = False
                    x2.Visible = True
                    v2 = "X"
                    ganoCpu = True
                ElseIf v1 = "X" And v2 = "X" And btn3.Visible = True
                    btn3.Visible = False
                    x3.Visible = True
                    v3 = "X"
                    ganoCpu = True
                ElseIf v5 = "X" And v7 = "X" And btn3.Visible = True
                    btn3.Visible = False
                    x3.Visible = True
                    v3 = "X"
                    ganoCpu = True
                ElseIf v6 = "X" And v9 = "X" And btn3.Visible = True
                    btn3.Visible = False
                    x3.Visible = True
                    v3 = "X"
                    ganoCpu = True
                ElseIf v5 = "X" And v6 = "X" And btn4.Visible = True
                    btn4.Visible = False
                    x4.Visible = True
                    v4 = "X"
                    ganoCpu = True
                ElseIf v1 = "X" And v7 = "X" And btn4.Visible = True
                    btn4.Visible = False
                    x4.Visible = True
                    v4 = "X"
                    ganoCpu = True
                ElseIf v2 = "X" And v8 = "X" And btn5.Visible = True
                    btn5.Visible = False
                    x5.Visible = True
                    v5 = "X"
                    ganoCpu = True
                ElseIf v1 = "X" And v9 = "X" And btn5.Visible = True
                    btn5.Visible = False
                    x5.Visible = True
                    v5 = "X"
                    ganoCpu = True
                ElseIf v3 = "X" And v7 = "X" And btn5.Visible = True
                    btn5.Visible = False
                    x5.Visible = True
                    v5 = "X"
                    ganoCpu = True
                ElseIf v4 = "X" And v6 = "X" And btn5.Visible = True
                    btn5.Visible = False
                    x5.Visible = True
                    v5 = "X"
                    ganoCpu = True
                ElseIf v4 = "X" And v5 = "X" And btn6.Visible = True
                    btn6.Visible = False
                    x6.Visible = True
                    v6 = "X"
                    ganoCpu = True
                ElseIf v3 = "X" And v9 = "X" And btn6.Visible = True
                    btn6.Visible = False
                    x6.Visible = True
                    v6 = "X"
                    ganoCpu = True
                ElseIf v1 = "X" And v4 = "X" And btn7.Visible = True
                    btn7.Visible = False
                    x7.Visible = True
                    v7 = "X"
                    ganoCpu = True
                ElseIf v3 = "X" And v5 = "X" And btn7.Visible = True
                    btn7.Visible = False
                    x7.Visible = True
                    v7 = "X"
                    ganoCpu = True
                ElseIf v8 = "X" And v9 = "X" And btn7.Visible = True
                    btn7.Visible = False
                    x7.Visible = True
                    v7 = "X"
                    ganoCpu = True
                ElseIf v7 = "X" And v9 = "X" And btn8.Visible = True
                    btn8.Visible = False
                    x8.Visible = True
                    v8 = "X"
                    ganoCpu = True
                ElseIf v2 = "X" And v5 = "X" And btn8.Visible = True
                    btn8.Visible = False
                    x8.Visible = True
                    v8 = "X"
                    ganoCpu = True
                ElseIf v3 = "X" And v6 = "X" And btn9.Visible = True
                    btn9.Visible = False
                    x9.Visible = True
                    v9 = "X"
                    ganoCpu = True
                ElseIf v1 = "X" And v5 = "X" And btn9.Visible = True
                    btn9.Visible = False
                    x9.Visible = True
                    v9 = "X"
                    ganoCpu = True
                ElseIf v7 = "X" And v8 = "X" And btn9.Visible = True
                    btn9.Visible = False
                    x9.Visible = True
                    v9 = "X"
                    ganoCpu = True
                End If
            End If

        ElseIf j2 = "CPU"
            If v2 = "0" And v3 = "0" And btn1.Visible = True Then
                btn1.Visible = False
                c1.Visible = True
                v1 = "0"
                ganoCpu = True
            ElseIf v5 = "0" And v9 = "0" And btn1.Visible = True
                btn1.Visible = False
                c1.Visible = True
                v1 = "0"
                ganoCpu = True
            ElseIf v4 = "0" And v7 = "0" And btn1.Visible = True
                btn1.Visible = False
                c1.Visible = True
                v1 = "0"
                ganoCpu = True
            ElseIf v1 = "0" And v3 = "0" And btn2.Visible = True
                btn2.Visible = False
                c2.Visible = True
                v2 = "0"
                ganoCpu = True
            ElseIf v5 = "0" And v8 = "0" And btn2.Visible = True
                btn2.Visible = False
                c2.Visible = True
                v2 = "0"
                ganoCpu = True
            ElseIf v1 = "0" And v2 = "0" And btn3.Visible = True
                btn3.Visible = False
                c3.Visible = True
                v3 = "0"
                ganoCpu = True
            ElseIf v5 = "0" And v7 = "0" And btn3.Visible = True
                btn3.Visible = False
                c3.Visible = True
                v3 = "0"
                ganoCpu = True
            ElseIf v6 = "0" And v9 = "0" And btn3.Visible = True
                btn3.Visible = False
                c3.Visible = True
                v3 = "0"
                ganoCpu = True
            ElseIf v5 = "0" And v6 = "0" And btn4.Visible = True
                btn4.Visible = False
                c4.Visible = True
                v4 = "0"
                ganoCpu = True
            ElseIf v1 = "0" And v7 = "0" And btn4.Visible = True
                btn4.Visible = False
                c4.Visible = True
                v4 = "0"
                ganoCpu = True
            ElseIf v2 = "0" And v8 = "0" And btn5.Visible = True
                btn5.Visible = False
                c5.Visible = True
                v5 = "0"
                ganoCpu = True
            ElseIf v1 = "0" And v9 = "0" And btn5.Visible = True
                btn5.Visible = False
                c5.Visible = True
                v5 = "0"
                ganoCpu = True
            ElseIf v3 = "0" And v7 = "0" And btn5.Visible = True
                btn5.Visible = False
                c5.Visible = True
                v5 = "0"
                ganoCpu = True
            ElseIf v4 = "0" And v6 = "0" And btn5.Visible = True
                btn5.Visible = False
                c5.Visible = True
                v5 = "0"
                ganoCpu = True
            ElseIf v4 = "0" And v5 = "0" And btn6.Visible = True
                btn6.Visible = False
                c6.Visible = True
                v6 = "0"
                ganoCpu = True
            ElseIf v3 = "0" And v9 = "0" And btn6.Visible = True
                btn6.Visible = False
                c6.Visible = True
                v6 = "0"
                ganoCpu = True
            ElseIf v1 = "0" And v4 = "0" And btn7.Visible = True
                btn7.Visible = False
                c7.Visible = True
                v7 = "0"
                ganoCpu = True
            ElseIf v3 = "0" And v5 = "0" And btn7.Visible = True
                btn7.Visible = False
                c7.Visible = True
                v7 = "0"
                ganoCpu = True
            ElseIf v8 = "0" And v9 = "0" And btn7.Visible = True
                btn7.Visible = False
                c7.Visible = True
                v7 = "0"
                ganoCpu = True
            ElseIf v7 = "0" And v9 = "0" And btn8.Visible = True
                btn8.Visible = False
                c8.Visible = True
                v8 = "0"
                ganoCpu = True
            ElseIf v2 = "0" And v5 = "0" And btn8.Visible = True
                btn8.Visible = False
                c8.Visible = True
                v8 = "0"
                ganoCpu = True
            ElseIf v3 = "0" And v6 = "0" And btn9.Visible = True
                btn9.Visible = False
                c9.Visible = True
                v9 = "0"
                ganoCpu = True
            ElseIf v1 = "0" And v5 = "0" And btn9.Visible = True
                btn9.Visible = False
                c9.Visible = True
                v9 = "0"
                ganoCpu = True
            ElseIf v7 = "0" And v8 = "0" And btn9.Visible = True
                btn9.Visible = False
                c9.Visible = True
                v9 = "0"
                ganoCpu = True
            End If
        End If
    End Sub

    Private Sub defender()
        If j1 = "CPU" Then

            If v2 = "0" And v3 = "0" And btn1.Visible = True Then
                btn1.Visible = False
                x1.Visible = True
                v1 = "X"
                defenderCpu = True
            ElseIf v5 = "0" And v9 = "0" And btn1.Visible = True
                btn1.Visible = False
                x1.Visible = True
                v1 = "X"
                defenderCpu = True
            ElseIf v4 = "0" And v7 = "0" And btn1.Visible = True
                btn1.Visible = False
                x1.Visible = True
                v1 = "X"
                defenderCpu = True
                If v1 = "0" And v3 = "0" And btn2.Visible = True Then
                    btn2.Visible = False
                    x2.Visible = True
                    v2 = "X"
                    defenderCpu = True
                ElseIf v5 = "0" And v8 = "0" And btn2.Visible = True
                    btn2.Visible = False
                    x2.Visible = True
                    v2 = "X"
                    defenderCpu = True
                ElseIf v1 = "0" And v2 = "0" And btn3.Visible = True
                    btn3.Visible = False
                    x3.Visible = True
                    v3 = "X"
                    defenderCpu = True
                ElseIf v5 = "0" And v7 = "0" And btn3.Visible = True
                    btn3.Visible = False
                    x3.Visible = True
                    v3 = "X"
                    defenderCpu = True
                ElseIf v6 = "0" And v9 = "0" And btn3.Visible = True
                    btn3.Visible = False
                    x3.Visible = True
                    v3 = "X"
                    defenderCpu = True
                ElseIf v5 = "0" And v6 = "0" And btn4.Visible = True
                    btn4.Visible = False
                    x4.Visible = True
                    v4 = "X"
                    defenderCpu = True
                ElseIf v1 = "0" And v7 = "0" And btn4.Visible = True
                    btn4.Visible = False
                    x4.Visible = True
                    v4 = "X"
                    defenderCpu = True
                ElseIf v2 = "0" And v8 = "0" And btn5.Visible = True
                    btn5.Visible = False
                    x5.Visible = True
                    v5 = "X"
                    defenderCpu = True
                ElseIf v1 = "0" And v9 = "0" And btn5.Visible = True
                    btn5.Visible = False
                    x5.Visible = True
                    v5 = "X"
                    defenderCpu = True
                ElseIf v3 = "0" And v7 = "0" And btn5.Visible = True
                    btn5.Visible = False
                    x5.Visible = True
                    v5 = "X"
                    defenderCpu = True
                ElseIf v4 = "0" And v6 = "0" And btn5.Visible = True
                    btn5.Visible = False
                    x5.Visible = True
                    v5 = "X"
                    defenderCpu = True
                ElseIf v4 = "0" And v5 = "0" And btn6.Visible = True
                    btn6.Visible = False
                    x6.Visible = True
                    v6 = "X"
                    defenderCpu = True
                ElseIf v3 = "0" And v9 = "0" And btn6.Visible = True
                    btn6.Visible = False
                    x6.Visible = True
                    v6 = "X"
                    defenderCpu = True
                ElseIf v1 = "0" And v4 = "0" And btn7.Visible = True
                    btn7.Visible = False
                    x7.Visible = True
                    v7 = "X"
                    defenderCpu = True
                ElseIf v3 = "0" And v5 = "0" And btn7.Visible = True
                    btn7.Visible = False
                    x7.Visible = True
                    v7 = "X"
                    defenderCpu = True
                ElseIf v8 = "0" And v9 = "0" And btn7.Visible = True
                    btn7.Visible = False
                    x7.Visible = True
                    v7 = "X"
                    defenderCpu = True
                ElseIf v7 = "0" And v9 = "0" And btn8.Visible = True
                    btn8.Visible = False
                    x8.Visible = True
                    v8 = "X"
                    defenderCpu = True
                ElseIf v2 = "0" And v5 = "0" And btn8.Visible = True
                    btn8.Visible = False
                    x8.Visible = True
                    v8 = "X"
                    defenderCpu = True
                ElseIf v3 = "0" And v6 = "0" And btn9.Visible = True
                    btn9.Visible = False
                    x9.Visible = True
                    v9 = "X"
                    defenderCpu = True
                ElseIf v1 = "0" And v5 = "0" And btn9.Visible = True
                    btn9.Visible = False
                    x9.Visible = True
                    v9 = "X"
                    defenderCpu = True
                ElseIf v7 = "0" And v8 = "0" And btn9.Visible = True
                    btn9.Visible = False
                    x9.Visible = True
                    v9 = "X"
                    defenderCpu = True
                End If
            End If

        ElseIf j2 = "CPU"
            If v2 = "X" And v3 = "X" And btn1.Visible = True Then
                btn1.Visible = False
                c1.Visible = True
                v1 = "0"
                defenderCpu = True
            ElseIf v5 = "X" And v9 = "X" And btn1.Visible = True
                btn1.Visible = False
                c1.Visible = True
                v1 = "0"
                defenderCpu = True
            ElseIf v4 = "X" And v7 = "X" And btn1.Visible = True
                btn1.Visible = False
                c1.Visible = True
                v1 = "0"
                defenderCpu = True
            ElseIf v1 = "X" And v3 = "X" And btn2.Visible = True
                btn2.Visible = False
                c2.Visible = True
                v2 = "0"
                defenderCpu = True
            ElseIf v5 = "X" And v8 = "X" And btn2.Visible = True
                btn2.Visible = False
                c2.Visible = True
                v2 = "0"
                defenderCpu = True
            ElseIf v1 = "X" And v2 = "X" And btn3.Visible = True
                btn3.Visible = False
                c3.Visible = True
                v3 = "0"
                defenderCpu = True
            ElseIf v5 = "X" And v7 = "X" And btn3.Visible = True
                btn3.Visible = False
                c3.Visible = True
                v3 = "0"
                defenderCpu = True
            ElseIf v6 = "X" And v9 = "X" And btn3.Visible = True
                btn3.Visible = False
                c3.Visible = True
                v3 = "0"
                defenderCpu = True
            ElseIf v5 = "X" And v6 = "X" And btn4.Visible = True
                btn4.Visible = False
                c4.Visible = True
                v4 = "0"
                defenderCpu = True
            ElseIf v1 = "X" And v7 = "X" And btn4.Visible = True
                btn4.Visible = False
                c4.Visible = True
                v4 = "0"
                defenderCpu = True
            ElseIf v2 = "X" And v8 = "X" And btn5.Visible = True
                btn5.Visible = False
                c5.Visible = True
                v5 = "0"
                defenderCpu = True
            ElseIf v1 = "X" And v9 = "X" And btn5.Visible = True
                btn5.Visible = False
                c5.Visible = True
                v5 = "0"
                defenderCpu = True
            ElseIf v3 = "X" And v7 = "X" And btn5.Visible = True
                btn5.Visible = False
                c5.Visible = True
                v5 = "0"
                defenderCpu = True
            ElseIf v4 = "X" And v6 = "X" And btn5.Visible = True
                btn5.Visible = False
                c5.Visible = True
                v5 = "0"
                defenderCpu = True
            ElseIf v4 = "X" And v5 = "X" And btn6.Visible = True
                btn6.Visible = False
                c6.Visible = True
                v6 = "0"
                defenderCpu = True
            ElseIf v3 = "X" And v9 = "X" And btn6.Visible = True
                btn6.Visible = False
                c6.Visible = True
                v6 = "0"
                defenderCpu = True
            ElseIf v1 = "X" And v4 = "X" And btn7.Visible = True
                btn7.Visible = False
                c7.Visible = True
                v7 = "0"
                defenderCpu = True
            ElseIf v3 = "X" And v5 = "X" And btn7.Visible = True
                btn7.Visible = False
                c7.Visible = True
                v7 = "0"
                defenderCpu = True
            ElseIf v8 = "X" And v9 = "X" And btn7.Visible = True
                btn7.Visible = False
                c7.Visible = True
                v7 = "0"
                defenderCpu = True
            ElseIf v7 = "X" And v9 = "X" And btn8.Visible = True
                btn8.Visible = False
                c8.Visible = True
                v8 = "0"
                defenderCpu = True
            ElseIf v2 = "X" And v5 = "X" And btn8.Visible = True
                btn8.Visible = False
                c8.Visible = True
                v8 = "0"
                defenderCpu = True
            ElseIf v3 = "X" And v6 = "X" And btn9.Visible = True
                btn9.Visible = False
                c9.Visible = True
                v9 = "0"
                defenderCpu = True
            ElseIf v1 = "X" And v5 = "X" And btn9.Visible = True
                btn9.Visible = False
                c9.Visible = True
                v9 = "0"
                defenderCpu = True
            ElseIf v7 = "X" And v8 = "X" And btn9.Visible = True
                btn9.Visible = False
                c9.Visible = True
                v9 = "0"
                defenderCpu = True
            End If
        End If

        If defenderCpu = True Then
            checarGanador()
            cambiarTurno()
        End If
    End Sub
End Class
