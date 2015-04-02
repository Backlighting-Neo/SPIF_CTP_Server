Public Class SPIF_MetaData
    Structure Item
        Dim JJC As Double
        Dim MMC As Double
        Dim KPC As Double
    End Structure

    Public MetaData(3) As Item

    Public Sub SetData(ByVal ParamArray DataArray() As Double)
        Me.MetaData(2).JJC = DataArray(0)
        Me.MetaData(1).JJC = DataArray(1)
        Me.MetaData(0).JJC = DataArray(2)
        Me.MetaData(2).MMC = DataArray(3)
        Me.MetaData(1).MMC = DataArray(4)
        Me.MetaData(0).MMC = DataArray(5)
        Me.MetaData(2).KPC = DataArray(6)
        Me.MetaData(1).KPC = DataArray(7)
        Me.MetaData(0).KPC = DataArray(8)
        Me.MetaData(3).JJC = Me.MetaData(0).JJC + Me.MetaData(1).JJC + Me.MetaData(2).JJC
        Me.MetaData(3).MMC = Me.MetaData(0).MMC + Me.MetaData(1).MMC + Me.MetaData(2).MMC
        Me.MetaData(3).KPC = Me.MetaData(0).KPC + Me.MetaData(1).KPC + Me.MetaData(2).KPC
    End Sub
End Class
