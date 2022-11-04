Public Class Form1
    Function CheckRows(data, pencil, ByRef changed)
        'Looks along the row to discard numbers in it
        For row As Integer = 0 To 8
            For col As Integer = 0 To 8
                For alongRow As Integer = 0 To 8
                    If data(row, col) = False Then
                        If data(row, alongRow) AndAlso pencil(row, col)(data(row, alongRow) - 1) = True Then
                            pencil(row, col)(data(row, alongRow) - 1) = False
                            changed = True
                        End If
                    End If
                Next
            Next
        Next

        Return pencil
    End Function

    Function CheckCols(data, pencil, ByRef changed)

        'Looks down column to discard numbers in it
        For row As Integer = 0 To 8
            For col As Integer = 0 To 8
                For downCol As Integer = 0 To 8
                    If data(row, col) = False Then
                        If data(downCol, col) AndAlso pencil(row, col)(data(downCol, col) - 1) = True Then
                            pencil(row, col)(data(downCol, col) - 1) = False
                            changed = True
                        End If
                    End If
                Next
            Next
        Next


        Return pencil
    End Function

    Function CheckQuads(data, pencil, quadrant, ByRef changed)
        For row As Integer = 0 To 8
            For col As Integer = 0 To 8
                If data(row, col) = False Then
                    For row2 As Integer = 0 To 8
                        For col2 As Integer = 0 To 8
                            If data(row2, col2) AndAlso quadrant(row, col) = quadrant(row2, col2) AndAlso pencil(row, col)(data(row2, col2) - 1) Then
                                pencil(row, col)(data(row2, col2) - 1) = False
                                changed = True
                            End If
                        Next
                    Next
                End If
            Next
        Next

        Return pencil
    End Function


    'Next up - make this check the rows and cols for where there's only one pencil mark so that must be it
    Sub Consolidate(data, pencil, quadrant, ByRef changed)

        For row As Integer = 0 To 8
            For col As Integer = 0 To 8
                If data(row, col) = False Then

                    Dim count As Integer = 0
                    For item As Integer = 0 To 8
                        If pencil(row, col)(item) = True Then
                            count += 1
                        End If
                    Next

                    If count = 1 Then

                        For item As Integer = 0 To 8
                            If pencil(row, col)(item) = True Then
                                ChangeLabel(row, col, item + 1)
                            End If
                        Next

                    ElseIf count = 0 Then
                        Button1.Text = "Unsolvable"
                    End If
                End If
            Next
        Next


        'Looks at rows and sees if there's only one box each number could go in
        For row As Integer = 0 To 8
            For col As Integer = 0 To 8
                If data(row, col) = False Then

                    Dim unique(8) As Boolean
                    For index As Integer = 0 To 8
                        If pencil(row, col)(index) = True Then
                            unique(index) = True
                        Else
                            unique(index) = False
                        End If
                    Next

                    'Looks at rows and sees if there's only one box each number could go in
                    For col2 As Integer = 0 To 8
                        If col <> col2 Then
                            If data(row, col2) = False Then
                                For index As Integer = 0 To 8
                                    If pencil(row, col)(index) = True AndAlso pencil(row, col2)(index) = True Then
                                        unique(index) = False
                                    End If
                                Next
                            End If
                        End If
                    Next

                    For item As Integer = 0 To 8
                        If unique(item) = True Then
                            ChangeLabel(row, col, item + 1)
                            changed = True
                        End If
                    Next


                    For index As Integer = 0 To 8
                        If pencil(row, col)(index) = True Then
                            unique(index) = True
                        Else
                            unique(index) = False
                        End If
                    Next

                    'Looks at columns and sees if it's unique solution
                    For row2 As Integer = 0 To 8
                        If row <> row2 Then
                            If data(row2, col) = False Then
                                For index As Integer = 0 To 8
                                    If pencil(row, col)(index) = True AndAlso pencil(row2, col)(index) = True Then
                                        unique(index) = False
                                    End If
                                Next
                            End If
                        End If
                    Next

                    For item As Integer = 0 To 8
                        If unique(item) = True Then
                            ChangeLabel(row, col, item + 1)
                            changed = True
                        End If
                    Next


                    For index As Integer = 0 To 8
                        If pencil(row, col)(index) = True Then
                            unique(index) = True
                        Else
                            unique(index) = False
                        End If
                    Next

                    'Looks at quadrant and sees if it's unique
                    For row2 As Integer = 0 To 8
                        For col2 As Integer = 0 To 8
                            If row2 <> row OrElse col2 <> col Then
                                If quadrant(row, col) = quadrant(row2, col2) AndAlso data(row2, col2) = False Then
                                    For index As Integer = 0 To 8
                                        If pencil(row, col)(index) = True AndAlso pencil(row2, col2)(index) = True Then
                                            unique(index) = False
                                        End If
                                    Next
                                End If
                            End If
                        Next
                    Next

                    For item As Integer = 0 To 8
                        If unique(item) = True Then
                            ChangeLabel(row, col, item + 1)
                            changed = True
                        End If
                    Next
                End If
            Next
        Next


    End Sub



    Function InitPencil(data)
        Dim pencil(8, 8) As List(Of Boolean)

        'I have made an array of lists of booleans ok, just wait it should work eventually
        For row As Integer = 0 To 8
            For col As Integer = 0 To 8
                If data(row, col) = False Then
                    'Sets the pencil for each number as True, as it's currently possible
                    pencil(row, col) = New List(Of Boolean)({True, True, True, True, True, True, True, True, True})
                End If
            Next
        Next

        Return pencil
    End Function

    Function FindQuadrant(data)
        Dim quadrant(8, 8) As Integer

        'Assigns a quadrant to each field
        For row = 0 To 8
            For col = 0 To 8
                Select Case row
                    Case 0 To 2
                        Select Case col
                            Case 0 To 2
                                quadrant(row, col) = 1
                            Case 3 To 5
                                quadrant(row, col) = 2
                            Case 6 To 8
                                quadrant(row, col) = 3
                        End Select

                    Case 3 To 5
                        Select Case col
                            Case 0 To 2
                                quadrant(row, col) = 4
                            Case 3 To 5
                                quadrant(row, col) = 5
                            Case 6 To 8
                                quadrant(row, col) = 6
                        End Select

                    Case 6 To 8
                        Select Case col
                            Case 0 To 2
                                quadrant(row, col) = 7
                            Case 3 To 5
                                quadrant(row, col) = 8
                            Case 6 To 8
                                quadrant(row, col) = 9
                        End Select
                End Select
            Next
        Next


        Return quadrant
    End Function

    Function GetData()
        Dim data(8, 8) As Integer
        Dim labelTarget As TextBox = Nothing
        Dim ref As Integer

        For row As Integer = 0 To 8
            For col As Integer = 0 To 8
                ref = GetRef(row, col)
                labelTarget = CType(SodukuGrid.Controls("TextBox" & ref), TextBox)

                'Gets the data from the grid, as I haven't passed it to the button as I don't get vb that well
                Try
                    data(row, col) = CType(labelTarget.Text, String)
                Catch ex As Exception
                    data(row, col) = Nothing
                End Try

            Next
        Next


        Return data
    End Function
    Function GetRef(row, col)
        Dim ref As Integer
        row += 1
        col += 1

        ref = ((row - 1) * 9) + col
        Return ref
    End Function

    Sub ChangeLabel(row, col, text)
        Dim ref As Integer
        Dim labelTarget As TextBox = Nothing

        ref = GetRef(row, col)
        labelTarget = CType(SodukuGrid.Controls("TextBox" & ref), TextBox)
        labelTarget.Text = text

    End Sub

    Function Initialise()
        Dim ref As Integer
        Dim data(,) As Integer = {
        {Nothing, 5, 1, Nothing, 9, Nothing, 2, Nothing, Nothing},
        {Nothing, 8, 4, 2, Nothing, Nothing, Nothing, 5, Nothing},
        {7, 2, Nothing, 8, Nothing, 5, Nothing, 4, 1},
        {8, Nothing, Nothing, Nothing, Nothing, 2, Nothing, Nothing, 5},
        {Nothing, 7, Nothing, Nothing, Nothing, Nothing, Nothing, 3, Nothing},
        {2, Nothing, Nothing, 6, Nothing, Nothing, Nothing, Nothing, 7},
        {4, 1, Nothing, 5, Nothing, 9, Nothing, 6, 3},
        {Nothing, 3, Nothing, Nothing, Nothing, 6, 8, 2, Nothing},
        {Nothing, Nothing, 2, Nothing, 7, Nothing, 5, 1, Nothing}
}


        For row As Integer = 0 To 8
            For col As Integer = 0 To 8
                Dim textBoxTarget As TextBox
                ref = GetRef(row, col)
                textBoxTarget = CType(SodukuGrid.Controls("TextBox" & ref), TextBox)
                textBoxTarget.Anchor = AnchorStyles.None
                textBoxTarget.BorderStyle = BorderStyle.None
            Next
        Next




        For row As Integer = 0 To 8
            For col As Integer = 0 To 8
                Dim labelTarget As TextBox = Nothing
                ref = GetRef(row, col)
                labelTarget = CType(SodukuGrid.Controls("TextBox" & ref), TextBox)
                If data(row, col) Then
                    labelTarget.Text = data(row, col)
                Else
                    labelTarget.Text = ""
                End If
            Next
        Next

        Return data
    End Function

    Sub Solve(data, ByRef changed, ByRef pencil)


        Dim quadrant(8, 8) As Integer


        quadrant = FindQuadrant(data)



        pencil = CheckRows(data, pencil, changed)


        pencil = CheckCols(data, pencil, changed)
        pencil = CheckQuads(data, pencil, quadrant, changed)


        Consolidate(data, pencil, quadrant, changed)

    End Sub



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim data As Integer(,)
        data = Initialise()
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim data(,) As Integer
        Dim changed As Boolean = True
        Dim pencil(8, 8) As List(Of Boolean)


        data = GetData()
        pencil = InitPencil(data)

        While changed = True
            changed = False
            data = GetData()
            Solve(data, changed, pencil)

        End While

    End Sub
End Class

