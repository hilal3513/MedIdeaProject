Sub Main
Dim oldDB(eskiveritabani) As Object
Dim oldTable(eskitablo)As Object
Dim oldRS(eskikayitlar) As Object
Dim oldRec(eskikayit) As Object
Dim newDB(yeniveritabani As Object
Dim newTable (yenitablo)As Object
Dim newRec(yenikayit) As Object
Dim newRS(yenikayitlar) As Object
Set eskiveritabani = Client.Currentdatabase
Set eskitablo = eskiveritabani.TableDef
Set eskikayitlar = eskiveritabani.RecordSet
Set eskikayit = eskikayitlar.ActiveRecord
Set yenitablo = Client.NewTableDef
yenitablo.CopyFrom (eskitablo)
Set yeniveritabani = Client.NewDatabase (hilalkaratas.IMD ," ", yenitablo)
Set yenikayitlar = yeniveritabani.RecordSet
eskikayitlar.GetAt 72
deger = eskikyt.getcharvalue ("COL1", "0")
deger = eskikyt.getcharvalueat ("COL2", "0")
eskikayitlar.ToFirst
For i = 4 To eskikayitlar.Count()
eskikayitlar.Next
yenikayitlar.AppendRecord (yenikayit)
Next i
yenikayitlar.ReleaseDatabase
eskikayitlar.ReleaseDatabase
Set yenikayitlar = Nothing
Set eskikayitlar = Nothing
End Sub
