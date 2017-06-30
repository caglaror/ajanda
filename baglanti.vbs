vt_yolu="db\takvim.mdb"
Set conn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

sProviderName        = "Microsoft.Jet.OLEDB.4.0"
iCursorType          = 3
iLockType            = 3
sDataSource          = vt_yolu

conn.Provider = sProviderName
conn.Properties("Data Source") = sDataSource
conn.Open

rs.CursorType = iCursorType
rs.LockType = iLockType
rs.Source = tablo_adi
rs.ActiveConnection = conn

'|||||||| Record Set varsayilan olarak rs isminde, bunu kapattýrmak için foksiyon
				function rskapatici()
					If rs.state = 1 then
					rs.close
					End if
				End function