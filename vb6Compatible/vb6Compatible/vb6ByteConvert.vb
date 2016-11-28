Imports System
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Security.Permissions
Public Enum copyMode
    toByte = 0
    fromByte = 1
End Enum
Public Enum byteArea
    autoAllocation = 0
    fixed = 1
End Enum


Public Class ByteConvert

    Private Shared Function itembegin(ByRef byteData() As Byte, mode As copyMode, lengthByte As Integer, startidx As Integer, byteArea As byteArea, ByRef bytedataCount As Integer) As Boolean
        itembegin = False
        If lengthByte <= 0 Then Exit Function
        If startidx < 0 Then Exit Function
        Select Case mode
            Case copyMode.toByte
                Select Case byteArea
                    Case byteArea.fixed
                        bytedataCount = UBound(byteData)

                        If startidx > bytedataCount Then Exit Function
                        If startidx + lengthByte - 1 > bytedataCount Then Exit Function
                    Case byteArea.autoAllocation
                        If IsArray(byteData) = False Then
                            ReDim byteData(startidx + lengthByte - 1)
                        End If
                        bytedataCount = UBound(byteData)
                        If startidx + lengthByte - 1 > bytedataCount Then
                            bytedataCount = startidx + lengthByte - 1
                            ReDim Preserve byteData(bytedataCount)
                        End If
                End Select
            Case copyMode.fromByte
                bytedataCount = UBound(byteData)

                If startidx > bytedataCount Then Exit Function
                If startidx + lengthByte - 1 > bytedataCount Then Exit Function
        End Select

        itembegin = True
    End Function
    Shared Sub item(ByRef Data As String,
                         ByRef byteData() As Byte,
                         mode As copyMode,
                         ByVal lengthByte As Integer,
                         ByRef startidx As Integer,
                         Optional ByRef nextidx As Integer = 0,
                         Optional byteArea As byteArea = byteArea.autoAllocation)
        Dim arr() As Byte
        Dim arrCount As Integer
        Dim i As Integer
        Dim sp() As Byte
        Dim idx As Integer
        Dim byteDataCount As Integer
        If itembegin(byteData, mode, lengthByte, startidx, byteArea, byteDataCount) = False Then
            Exit Sub
        End If
        Select Case mode
            Case copyMode.toByte
                arr = System.Text.Encoding.GetEncoding(932).GetBytes(Data)
                arrCount = UBound(arr)
                'lengthByteに満たない場合、スペースを補填
                If lengthByte - 1 > arrCount Then
                    ReDim Preserve arr(lengthByte - 1)
                    sp = System.Text.Encoding.GetEncoding(932).GetBytes(" ")
                    For i = arrCount + 1 To lengthByte - 1
                        arr(i) = sp(0)
                    Next
                End If
                'lengthByteを超える文字は破棄する
                For i = 0 To lengthByte - 1
                    idx = i + startidx
                    If idx > byteDataCount Then Exit For
                    byteData(idx) = arr(i)
                Next
                nextidx = startidx + lengthByte
            Case copyMode.fromByte
                ReDim arr(lengthByte - 1)
                For i = 0 To lengthByte - 1
                    idx = startidx + i
                    If idx > byteDataCount Then Exit For
                    arr(i) = byteData(idx)
                Next
                Data = System.Text.Encoding.GetEncoding(932).GetString(arr)
                nextidx = startidx + lengthByte
        End Select

        Erase arr
        'Erase sp

    End Sub

    Shared Sub item(ByRef Data As Short,
                         ByRef byteData() As Byte,
                         mode As copyMode,
                         ByRef startidx As Integer,
                         Optional ByRef nextidx As Integer = 0,
                         Optional byteArea As byteArea = byteArea.autoAllocation,
                         Optional changeStartidxDataAlignment As Boolean = True)
        Dim arr() As Byte
        Dim i As Integer
        Dim byteDataCount As Integer

        If changeStartidxDataAlignment Then
            If startidx Mod Marshal.SizeOf(Data) Then
                startidx = (startidx \ Marshal.SizeOf(Data) + 1) * Marshal.SizeOf(Data)
            End If
        End If

        If itembegin(byteData, mode, Marshal.SizeOf(Data), startidx, byteArea, byteDataCount) = False Then
            Exit Sub
        End If
        Select Case mode
            Case copyMode.toByte
                arr = BitConverter.GetBytes(Data)
                For i = 0 To UBound(arr)
                    byteData(i + startidx) = arr(i)
                Next
                nextidx = startidx + UBound(arr) + 1
            Case copyMode.fromByte
                Data = BitConverter.ToInt16(byteData, startidx)
                nextidx = startidx + Marshal.SizeOf(Data)
        End Select

        Erase arr
    End Sub

    Shared Sub item(ByRef Data As Long,
                         ByRef byteData() As Byte,
                         mode As copyMode,
                         ByRef startidx As Integer,
                         Optional ByRef nextidx As Integer = 0,
                         Optional byteArea As byteArea = byteArea.autoAllocation,
                         Optional changeStartidxDataAlignment As Boolean = True)
        Dim arr() As Byte
        Dim i As Integer
        Dim byteDataCount As Integer
        If changeStartidxDataAlignment Then
            If startidx Mod Marshal.SizeOf(Data) Then
                startidx = (startidx \ Marshal.SizeOf(Data) + 1) * Marshal.SizeOf(Data)
            End If
        End If

        If itembegin(byteData, mode, Marshal.SizeOf(Data), startidx, byteArea, byteDataCount) = False Then
            Exit Sub
        End If
        Select Case mode
            Case copyMode.toByte
                arr = BitConverter.GetBytes(Data)
                For i = 0 To UBound(arr)
                    byteData(i + startidx) = arr(i)
                Next
                nextidx = startidx + UBound(arr) + 1
            Case copyMode.fromByte
                Data = BitConverter.ToInt64(byteData, startidx)
                nextidx = startidx + Marshal.SizeOf(Data)
        End Select

        Erase arr
    End Sub

    Shared Sub item(ByRef Data As Integer,
                         ByRef byteData() As Byte,
                         mode As copyMode,
                         ByRef startidx As Integer,
                         Optional ByRef nextidx As Integer = 0,
                         Optional byteArea As byteArea = byteArea.autoAllocation,
                         Optional changeStartidxDataAlignment As Boolean = True)
        Dim arr() As Byte
        Dim i As Integer
        Dim byteDataCount As Integer
        If changeStartidxDataAlignment Then
            If startidx Mod Marshal.SizeOf(Data) Then
                startidx = (startidx \ Marshal.SizeOf(Data) + 1) * Marshal.SizeOf(Data)
            End If
        End If

        If itembegin(byteData, mode, Marshal.SizeOf(Data), startidx, byteArea, byteDataCount) = False Then
            Exit Sub
        End If
        Select Case mode
            Case copyMode.toByte
                arr = BitConverter.GetBytes(Data)
                For i = 0 To UBound(arr)
                    byteData(i + startidx) = arr(i)
                Next
                nextidx = startidx + UBound(arr) + 1
            Case copyMode.fromByte
                Data = BitConverter.ToInt32(byteData, startidx)
                nextidx = startidx + Marshal.SizeOf(Data)
        End Select

        Erase arr
    End Sub

    Shared Sub item(ByRef Data As Single,
                         ByRef byteData() As Byte,
                         mode As copyMode,
                         ByRef startidx As Integer,
                         Optional ByRef nextidx As Integer = 0,
                         Optional byteArea As byteArea = byteArea.autoAllocation,
                         Optional changeStartidxDataAlignment As Boolean = True)
        Dim arr() As Byte
        Dim i As Integer
        If changeStartidxDataAlignment Then
            If startidx Mod Marshal.SizeOf(Data) Then
                startidx = (startidx \ Marshal.SizeOf(Data) + 1) * Marshal.SizeOf(Data)
            End If
        End If

        Dim byteDataCount As Integer
        If itembegin(byteData, mode, Marshal.SizeOf(Data), startidx, byteArea, byteDataCount) = False Then
            Exit Sub
        End If
        Select Case mode
            Case copyMode.toByte
                arr = BitConverter.GetBytes(Data)
                For i = 0 To UBound(arr)
                    byteData(i + startidx) = arr(i)
                Next
                nextidx = startidx + UBound(arr) + 1
            Case copyMode.fromByte
                Data = BitConverter.ToSingle(byteData, startidx)
                nextidx = startidx + Marshal.SizeOf(Data)
        End Select

        Erase arr
    End Sub

    Shared Sub item(ByRef Data As Double,
                         ByRef byteData() As Byte,
                         mode As copyMode,
                         ByRef startidx As Integer,
                         Optional ByRef nextidx As Integer = 0,
                         Optional byteArea As byteArea = byteArea.autoAllocation,
                         Optional changeStartidxDataAlignment As Boolean = True)
        Dim arr() As Byte
        Dim i As Integer
        Dim byteDataCount As Integer
        If changeStartidxDataAlignment Then
            If startidx Mod Marshal.SizeOf(Data) Then
                startidx = (startidx \ Marshal.SizeOf(Data) + 1) * Marshal.SizeOf(Data)
            End If
        End If

        If itembegin(byteData, mode, Marshal.SizeOf(Data), startidx, byteArea, byteDataCount) = False Then
            Exit Sub
        End If
        Select Case mode
            Case copyMode.toByte
                arr = BitConverter.GetBytes(Data)
                For i = 0 To UBound(arr)
                    byteData(i + startidx) = arr(i)
                Next
                nextidx = startidx + UBound(arr) + 1
            Case copyMode.fromByte
                Data = BitConverter.ToDouble(byteData, startidx)
                nextidx = startidx + Marshal.SizeOf(Data)
        End Select

        Erase arr
    End Sub
    Shared Sub item(ByRef Data As Decimal,
                         ByRef byteData() As Byte,
                         mode As copyMode,
                         ByRef startidx As Integer,
                         Optional ByRef nextidx As Integer = 0,
                         Optional byteArea As byteArea = byteArea.autoAllocation,
                         Optional changeStartidxDataAlignment As Boolean = True)
        Dim arr() As Byte
        Dim i As Integer
        Dim wdat As Int64
        Dim rdat As Decimal
        If changeStartidxDataAlignment Then
            If startidx Mod Marshal.SizeOf(wdat) Then
                startidx = (startidx \ Marshal.SizeOf(wdat) + 1) * Marshal.SizeOf(wdat)
            End If
        End If

        Dim byteDataCount As Integer
        If itembegin(byteData, mode, Marshal.SizeOf(wdat), startidx, byteArea, byteDataCount) = False Then
            Exit Sub
        End If
        Select Case mode
            Case copyMode.toByte
                rdat = Data
                rdat = rdat * 10000
                wdat = rdat
                arr = BitConverter.GetBytes(wdat)
                For i = 0 To UBound(arr)
                    byteData(startidx + i) = arr(i)
                Next
                nextidx = startidx + UBound(arr) + 1
            Case copyMode.fromByte
                wdat = BitConverter.ToInt64(byteData, startidx)
                rdat = New Decimal(wdat)
                rdat = rdat / 10000
                Data = rdat
                nextidx = startidx + Marshal.SizeOf(wdat)
        End Select

        Erase arr
    End Sub

End Class
