Imports System.Net

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Actualizar()
    End Sub

    Private Sub Button1_MouseClick(sender As Object, e As MouseEventArgs) Handles Button1.MouseClick
        If TextBox2.Text IsNot Nothing Then
            Dim newPersona As New Persona
            Dim selPersona As New Persona
            selPersona.leerPersona(TextBox2.Text)
            If TextBox2.Text <> selPersona.NDNI1 Then
                newPersona.NDNI1 = TextBox2.Text
                newPersona.NOM = TextBox1.Text
                newPersona.insertarPersona()
                Actualizar()
                MessageBox.Show("Se ha añadido la persona " + newPersona.NDNI1 + " satisfactoriamente")
            End If
        End If
    End Sub

    Private Sub Button2_MouseClick(sender As Object, e As MouseEventArgs) Handles Button2.MouseClick
        If TextBox3.Text IsNot Nothing Then
            Dim newAsignatura As New Asignatura
            Dim selAsignatura As New Asignatura
            selAsignatura.leerAsignatura(TextBox3.Text)
            If TextBox3.Text <> selAsignatura.NOM Then
                newAsignatura.NOM = TextBox3.Text
                newAsignatura.Curso = TextBox5.Text
                newAsignatura.Horas = TextBox6.Text
                newAsignatura.insertarAsignatura()
                '  Actualizar()
                MessageBox.Show("Se ha añadido la Asignatura " + newAsignatura.NOM + " satisfactoriamente")
            End If
        End If
    End Sub

    Private Sub borrar_todo_Click(sender As Object, e As EventArgs) Handles borrar_todo.Click
        Dim newPersona As New Persona
        newPersona.eliminarBBDD()
        Actualizar()
    End Sub

    Private Sub salir_Click(sender As Object, e As EventArgs) Handles salir.Click
        Me.Close()
    End Sub

    Private Sub Actualizar()
        Lista_Personas.Items.Clear()
        Dim newPersona As New Persona
        Dim lista As Collection
        Dim i As Integer = 1
        lista = newPersona.leertodos()
        If lista IsNot Nothing Then
            While i <= lista.Count
                Lista_Personas.Items.Add(lista.Item(i).mDNI1)
                i = i + 1
            End While
        End If
    End Sub

End Class
