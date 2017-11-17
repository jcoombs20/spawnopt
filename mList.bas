Attribute VB_Name = "Module1"
Public mList1 As New Collection
Public mList2 As New Collection
Public mList3 As New Collection
Public mList4 As New Collection
Public mList5 As New Collection
Public mList6 As New Collection
Public mList7 As New Collection
Public mList8 As New Collection
Public mList9 As New Collection
Public mList10 As New Collection

Public fList4 As New Collection

Public Function getFList4()
    For a = 1 To 1
        For b = a + 1 To 4
            For c = a + 1 To 4
                If c <> b Then
                    For d = c + 1 To 4
                        If d <> b And d <> c Then
                            fList4.Add Array((a * 2) - 1, (a * 2), (b * 2) - 1, (b * 2), (c * 2) - 1, (c * 2), (d * 2) - 1, (d * 2))
                        End If
                    Next d
                End If
            Next c
        Next b
    Next a
End Function

Public Function getList1()
    For a = 1 To 1
        mList1.Add Array(a)
    Next a
End Function

Public Function getList2()
    For a = 1 To 2
        For b = 1 To 2
            If b <> a Then
                mList2.Add Array(a, b)
            End If
        Next b
    Next a
End Function

Public Function getList3()
    For a = 1 To 3
        For b = 1 To 3
            If b <> a Then
                For c = 1 To 3
                    If c <> a And c <> b Then
                        mList3.Add Array(a, b, c)
                    End If
                Next c
            End If
        Next b
    Next a
End Function

Public Function getList4()
    For a = 1 To 4
        For b = 1 To 4
            If b <> a Then
                For c = 1 To 4
                    If c <> a And c <> b Then
                        For d = 1 To 4
                            If d <> a And d <> b And d <> c Then
                                mList4.Add Array(a, b, c, d)
                            End If
                        Next d
                    End If
                Next c
            End If
        Next b
    Next a
End Function

Public Function getList5()
    For a = 1 To 5
        For b = 1 To 5
            If b <> a Then
                For c = 1 To 5
                    If c <> a And c <> b Then
                        For d = 1 To 5
                            If d <> a And d <> b And d <> c Then
                                For e = 1 To 5
                                    If e <> a And e <> b And e <> c And e <> d Then
                                        mList5.Add Array(a, b, c, d, e)
                                    End If
                                Next e
                            End If
                        Next d
                    End If
                Next c
            End If
        Next b
    Next a
End Function

Public Function getList6()
    For a = 1 To 6
        For b = 1 To 6
            If b <> a Then
                For c = 1 To 6
                    If c <> a And c <> b Then
                        For d = 1 To 6
                            If d <> a And d <> b And d <> c Then
                                For e = 1 To 6
                                    If e <> a And e <> b And e <> c And e <> d Then
                                        For f = 1 To 6
                                            If f <> a And f <> b And f <> c And f <> d And f <> e Then
                                                mList6.Add Array(a, b, c, d, e, f)
                                            End If
                                        Next f
                                    End If
                                Next e
                            End If
                        Next d
                    End If
                Next c
            End If
        Next b
    Next a
End Function

Public Function getList7()
    For a = 1 To 7
        For b = 1 To 7
            If b <> a Then
                For c = 1 To 7
                    If c <> a And c <> b Then
                        For d = 1 To 7
                            If d <> a And d <> b And d <> c Then
                                For e = 1 To 7
                                    If e <> a And e <> b And e <> c And e <> d Then
                                        For f = 1 To 7
                                            If f <> a And f <> b And f <> c And f <> d And f <> e Then
                                                For g = 1 To 7
                                                    If g <> a And g <> b And g <> c And g <> d And g <> e And g <> f Then
                                                        mList7.Add Array(a, b, c, d, e, f, g)
                                                    End If
                                                Next g
                                            End If
                                        Next f
                                    End If
                                Next e
                            End If
                        Next d
                    End If
                Next c
            End If
        Next b
    Next a
End Function

Public Function getList8()
    For a = 1 To 8
        For b = 1 To 8
            If b <> a Then
                For c = 1 To 8
                    If c <> a And c <> b Then
                        For d = 1 To 8
                            If d <> a And d <> b And d <> c Then
                                For e = 1 To 8
                                    If e <> a And e <> b And e <> c And e <> d Then
                                        For f = 1 To 8
                                            If f <> a And f <> b And f <> c And f <> d And f <> e Then
                                                For g = 1 To 8
                                                    If g <> a And g <> b And g <> c And g <> d And g <> e And g <> f Then
                                                        For h = 1 To 8
                                                            If h <> a And h <> b And h <> c And h <> d And h <> e And h <> f And h <> g Then
                                                                mList8.Add Array(a, b, c, d, e, f, g, h)
                                                            End If
                                                        Next h
                                                    End If
                                                Next g
                                            End If
                                        Next f
                                    End If
                                Next e
                            End If
                        Next d
                    End If
                Next c
            End If
        Next b
    Next a
End Function

Public Function getList9()
    For a = 1 To 9
        For b = 1 To 9
            If b <> a Then
                For c = 1 To 9
                    If c <> a And c <> b Then
                        For d = 1 To 9
                            If d <> a And d <> b And d <> c Then
                                For e = 1 To 9
                                    If e <> a And e <> b And e <> c And e <> d Then
                                        For f = 1 To 9
                                            If f <> a And f <> b And f <> c And f <> d And f <> e Then
                                                For g = 1 To 9
                                                    If g <> a And g <> b And g <> c And g <> d And g <> e And g <> f Then
                                                        For h = 1 To 9
                                                            If h <> a And h <> b And h <> c And h <> d And h <> e And h <> f And h <> g Then
                                                                For i = 1 To 9
                                                                    If i <> a And i <> b And i <> c And i <> d And i <> e And i <> f And i <> g And i <> h Then
                                                                        mList9.Add Array(a, b, c, d, e, f, g, h, i)
                                                                    End If
                                                                Next i
                                                            End If
                                                        Next h
                                                    End If
                                                Next g
                                            End If
                                        Next f
                                    End If
                                Next e
                            End If
                        Next d
                    End If
                Next c
            End If
        Next b
    Next a
End Function

Public Function getList10()
    For a = 1 To 10
        For b = 1 To 10
            If b <> a Then
                For c = 1 To 10
                    If c <> a And c <> b Then
                        For d = 1 To 10
                            If d <> a And d <> b And d <> c Then
                                For e = 1 To 10
                                    If e <> a And e <> b And e <> c And e <> d Then
                                        For f = 1 To 10
                                            If f <> a And f <> b And f <> c And f <> d And f <> e Then
                                                For g = 1 To 10
                                                    If g <> a And g <> b And g <> c And g <> d And g <> e And g <> f Then
                                                        For h = 1 To 10
                                                            If h <> a And h <> b And h <> c And h <> d And h <> e And h <> f And h <> g Then
                                                                For i = 1 To 10
                                                                    If i <> a And i <> b And i <> c And i <> d And i <> e And i <> f And i <> g And i <> h Then
                                                                        For j = 1 To 10
                                                                            If j <> a And j <> b And j <> c And j <> d And j <> e And j <> f And j <> g And j <> h And j <> i Then
                                                                                mList10.Add Array(a, b, c, d, e, f, g, h, i, j)
                                                                            End If
                                                                        Next j
                                                                    End If
                                                                Next i
                                                            End If
                                                        Next h
                                                    End If
                                                Next g
                                            End If
                                        Next f
                                    End If
                                Next e
                            End If
                        Next d
                    End If
                Next c
            End If
        Next b
    Next a
End Function

