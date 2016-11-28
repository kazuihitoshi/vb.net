Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports vb6Compatible

<TestClass()> Public Class UnitTest1

    <TestMethod()> Public Sub toByteJpn_SpaceAdd()
        Dim dat As String
        Dim bdat() As Byte
        Dim startidx As Integer
        Dim nextidx As Integer
        Dim resultbyte() As Byte
        '//1 日本語文字列コンバート  space add
        dat = "あいうえお"
        ReDim bdat(0)
        vb6Compatible.ByteConvert.item(dat, bdat, copyMode.toByte, 20, startidx, nextidx)
        resultbyte = {&H82, &HA0, &H82, &HA2, &H82, &HA4, &H82, &HA6, &H82, &HA8, &H20, &H20, &H20, &H20, &H20, &H20, &H20, &H20, &H20, &H20}
        CollectionAssert.AreEqual(bdat, resultbyte)
        Assert.AreEqual(startidx, 0)
        Assert.AreEqual(nextidx, 20)
    End Sub

    <TestMethod()> Public Sub toByteJpn_OverLength()
        Dim dat As String
        Dim bdat() As Byte
        Dim startidx As Integer
        Dim nextidx As Integer
        Dim resultbyte() As Byte
        '//1 日本語文字列コンバート  space add
        dat = "あいうえお"
        ReDim bdat(0)
        vb6Compatible.ByteConvert.item(dat, bdat, copyMode.toByte, 4, startidx, nextidx)
        resultbyte = {&H82, &HA0, &H82, &HA2}
        CollectionAssert.AreEqual(bdat, resultbyte)
        Assert.AreEqual(startidx, 0)
        Assert.AreEqual(nextidx, 4)
    End Sub

    <TestMethod()> Public Sub toByteJpn_EqualLength()
        Dim dat As String
        Dim bdat() As Byte
        Dim startidx As Integer
        Dim nextidx As Integer
        Dim resultbyte() As Byte
        '//1 日本語文字列コンバート  space add
        dat = "あい"
        ReDim bdat(0)
        vb6Compatible.ByteConvert.item(dat, bdat, copyMode.toByte, 4, startidx, nextidx)
        resultbyte = {&H82, &HA0, &H82, &HA2}
        CollectionAssert.AreEqual(bdat, resultbyte)
        Assert.AreEqual(startidx, 0)
        Assert.AreEqual(nextidx, 4)

    End Sub

    <TestMethod()> Public Sub fromByteJpn()
        Dim dat As String
        Dim bdat() As Byte
        Dim startidx As Integer
        Dim nextidx As Integer
        Dim resultbyte() As Byte
        '//1 日本語文字列コンバート  space add
        bdat = {&H82, &HA0, &H82, &HA2, &H82, &HA4, &H82, &HA6, &H82, &HA8, &H20, &H20, &H20, &H20, &H20, &H20, &H20, &H20, &H20, &H20}
        vb6Compatible.ByteConvert.item(dat, bdat, copyMode.fromByte, 10, startidx, nextidx)
        resultbyte = {&H82, &HA0, &H82, &HA2}
        Assert.AreEqual(dat, "あいうえお")
        Assert.AreEqual(startidx, 0)
        Assert.AreEqual(nextidx, 10)
    End Sub

    <TestMethod()> Public Sub toByteJpnLong()
        Dim dat As String
        Dim bdat() As Byte
        Dim startidx As Integer
        Dim nextidx As Integer
        Dim resultbyte() As Byte
        Dim ldat As Integer

        resultbyte = {&H82, &HA0, 0, 0, &HFF, &HFF, &HFF, &HFF}
        '//1 日本語文字列コンバート  space add
        'bdat = {&H82, &HA0, &H82, &HA2, &H82, &HA4, &H82, &HA6, &H82, &HA8, &H20, &H20, &H20, &H20, &H20, &H20, &H20, &H20, &H20, &H20}

        vb6Compatible.ByteConvert.item("あ", bdat, copyMode.toByte, 2, startidx, nextidx)

        ldat = -1

        startidx = nextidx
        vb6Compatible.ByteConvert.item(ldat, bdat, copyMode.toByte, startidx, nextidx)

        CollectionAssert.AreEqual(bdat, resultbyte)
        Assert.AreEqual(startidx, 4)
        Assert.AreEqual(nextidx, 8)
    End Sub

End Class