<%@ Page Language="VB" %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.OleDb" %> 
<%@ Import Namespace="System.Configuration" %>
<%@ Import Namespace="System.Web.Configuration" %>

<%
    'Inherits System.Windows.Forms.Form
    'Create ADO.NET objects.
    Dim myConn As SqlConnection
    Dim myCmd As SqlCommand
	Dim myCmdi As SqlCommand
    Dim myReader As SqlDataReader
    'results As String
	Dim a, aaa, jmla, jmldata as integer
	Dim perihal0 as string = ""
	Dim nama0 as string = ""
	Dim posisi0 as String = ""
	Dim jenisund0 as String = ""
	Dim tembusan0 as String = ""
	Dim pengirim0 as String = ""
	Dim tgl0 as String = ""
	Dim jenissurat0 as String = ""
	Dim tujuan0 as String = ""
	Dim rubrik0 as String = ""
	Dim namafile0 as String = ""

	'Dim pageno as Integer = 1
	Dim pageno as String
	Dim jmlpage as Integer = 1
	Dim pagerows as Integer = 20
	Dim rowsperpage as Integer = 20
	Dim hal as Integer = 1
	Dim b as Long = 0
	Dim c as Integer = 0
	Dim idx as Integer = 0
	Dim itop as Integer = 1
	Dim ii as Integer = 0
	Dim k as Integer
	Dim userid as Integer = 0
	Dim namauserx as String = ""
	Dim prenamafile as String = ""
	Dim namatemp as String = ""
	Dim jenisoutlet as String = ""
	Dim filePath as String = ""
	Dim kcx as String = ""
	Dim maxid as String = ""
	Dim maxu as Integer = 0

	Dim SQL0 As String
	Dim SQL1 As String
	Dim SQL2 As String
	Dim SQL3 As String
	Dim SQLjo As String
	
	Dim statuskc as Integer = 0
	Dim statuskcp as Integer = 0
	
	Dim user0 as String
	Dim jenis0 as String
	Dim idq As String = ""
	Dim namauserq As String = ""
	Dim statusq As String = ""
	Dim jabatanq As String = ""
	Dim klpq As String = ""
	Dim unitq As String = ""
	Dim posisiq As String = ""
	Dim suratq As String = ""
	Dim sekreq As String = ""
	Dim divq As String = ""
	Dim posisix as String = ""
	Dim tahun as integer = 0
	Dim bulan as integer = 0
	Dim tglhari as Integer = 0
	Dim filepath1 as String
	Dim filepath2 as String
	Dim folder1 as String
	Dim folder2 as String
	Dim maxnu as String = ""
	Dim divstr as String = ""
	Dim idku as Integer = 0
	Dim nomorx as String = ""
	Dim nomor2x as String = ""
	Dim nomor3x as String = ""
	Dim nomor4x as String = ""
	Dim nomor5x as String = ""
	Dim nomor6x as String = ""
	Dim namax as String = ""
	Dim jmlnegox as Long = 0
	Dim vendorx as String = ""
	Dim jenisadax as String = ""
	
	Dim mohonid as Integer = 0
	Dim pksid as Integer = 0
	Dim suratmemo as String = ""

	user0 = Session("username")
	idq = Session("userid")
	namauserq = Session("namauser")
	jabatanq = Session("jabatan")
	klpq = Session("klp")
	unitq = Session("unit")
	posisiq = Session("posisi")
	posisix = Session("posisistr")
	suratq = Session("surat")
	divq = Session("divisi")

	
	If NOT Request.QueryString("mohonid") is Nothing Then
		 mohonid = Val(Request.QueryString("mohonid"))
	End If
	
	If NOT Request.QueryString("pksid") is Nothing Then
		 pksid = Val(Request.QueryString("pksid"))
	End If
	
	If NOT Request.QueryString("tahun") is Nothing Then
		 tahun = Val(Request.QueryString("tahun"))
	End If
	
	If NOT Request.QueryString("jmlnego") is Nothing Then
		 jmlnegox = Val(Request.QueryString("jmlnego"))
	End If
	
	
	If NOT Request.QueryString("tgl") is Nothing Then
		tgl0 = Request.QueryString("tgl")
		if tgl0.length() > 5 Then
			Dim tglarr(3) as String 
			tglarr = tgl0.Split("-")
			tgl0 = tglarr(2) & "-" & tglarr(1) & "-" & tglarr(0)
		End If
	End If
	
'Waktu -> prefix namafile
	Dim s As String = DateTime.Now.ToString("yyyyMMdd")
	Dim tglinputq As String = DateTime.Now.ToString("yyyy-MM-dd")
	Dim jamsaiki As String = DateTime.Now.ToString("HH:mm:ss")
	Dim tahunstr as String = DateTime.Now.ToString("yyyy")
	Dim tahunq as Integer = Val(tahunstr)
	s = s & Hour(Now) & Minute(Now) & Second(Now)
	prenamafile = s
	

	divstr = "OPR"
	
'***************************************************************************************************************	

	aaa = 1
    Try
        If (aaa=1)
			'myConn = New SqlConnection("Data Source=DESKTOP-S7N8DF9;Initial Catalog=KCT;User ID=sa;Password=sa;")
%>
<!--#include file="koneksi.aspx"-->
<%				
			
			myCmd = myConn.CreateCommand
			
'==============================================================================
' Cari nama pengadaan di tabel pengadaan
'==============================================================================
			SQL3 = "SELECT * FROM pengadaan WHERE ID=" & mohonid
			myCmd.CommandText = SQL3
			myConn.Open()
			myReader = myCmd.ExecuteReader()			
			If myReader.HasRows Then
				While myReader.Read()
					namax = myReader("nama").ToString()
					jenisadax = myReader("jenispengadaan").ToString()
				End While
			End If
			myConn.Close()
'==============================================================================
' Jika a=0 -> BIKIN nomor
'==============================================================================
			SQL3 = "SELECT * FROM pengadaan_pks WHERE pengadaanID=" & mohonid
'response.write(SQL3)
			myCmd.CommandText = SQL3
			myConn.Open()
			myReader = myCmd.ExecuteReader()		
			If myReader.HasRows Then
				While myReader.Read()
					nomorx = myReader("nomor").ToString()
					nomor2x = myReader("nomor2").ToString()
					nomor3x = myReader("nomor3").ToString()
					nomor4x = myReader("nomor4").ToString()
					nomor5x = myReader("nomor5").ToString()
					nomor6x = myReader("nomor6").ToString()
				End While
			End If
			myConn.Close
			
			b = 0
			If nomorx.length() > 3 Then
				a = 1
				b = b + 1
			End If
			If nomor2x.length() > 3 Then
				a = 2
				b = b + 1
			End If
			If nomor3x.length() > 3 Then
				a = 3
				b = b + 1
			End If
			If nomor4x.length() > 3 Then
				a = 4
				b = b + 1
			End If
			If nomor5x.length() > 3 Then
				a = 5
				b = b + 1
			End If
			If nomor6x.length() > 3 Then
				a = 6
				b = b + 1
			End If
			
			
			If b < 6 Then
				maxu = 0
	'** Cek max value of nomor urut

				
				SQLjo = "SELECT MAX(nourut) AS maxnourut FROM [surat_pks] WHERE (nomor LIKE '%" & divstr & "%' AND tahun=" & tahun & ")"
				myCmd.CommandText = SQLjo
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				If myReader.HasRows Then
					While myReader.Read()
						maxnu = myReader("maxnourut").ToString()
						maxu = Val(maxnu)
					End While
				End If
				myConn.Close()
				maxu = maxu + 1
				
'=====================================================================================================================
				Dim nomorxx as String = ""
				Dim nomorzz as String = ""
				nomorxx = maxu & "/PKS/OPR/" & tahun
				nomorzz = nomorxx
				vendorx = ""
				SQL0 = "INSERT INTO surat_pks(pengadaanID, tglpks, namapengadaan, jenisada, nourut, nomor, divisi, tahun, nominal, " & _
						"vendor, createduser, createddate) VALUES(" & mohonid & ",GETDATE(),'" & namax & "','" & jenisadax & _
						"'," & maxu & ",'" & nomorxx & "',1," & tahun & "," & jmlnegox & ",'" & vendorx & "'," & userid & ",GETDATE())"
		
				myCmd.CommandText = SQL0
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				myConn.Close()
'=====================================================================================================================
				SQL2 = "SELECT * FROM surat_pks WHERE (nomor='" & nomorxx & "' AND tahun=" & tahun & ")"
				myCmd.CommandText = SQL2
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				If myReader.HasRows Then
					While myReader.Read()
						idku = myReader("ID").ToString()
					End While
				End If
				myConn.Close()
'=====================================================================================================================
' Update nomor pada pengadaan pks
'=====================================================================================================================
				
				If a = 1 Then
					SQL3 = "UPDATE pengadaan_pks SET nomor2='" & nomorxx & "' WHERE ID=" & pksid
				Else If a = 2 Then
					SQL3 = "UPDATE pengadaan_pks SET nomor3='" & nomorxx & "' WHERE ID=" & pksid
				Else If a = 3 Then
					SQL3 = "UPDATE pengadaan_pks SET nomor4='" & nomorxx & "' WHERE ID=" & pksid
				Else If a = 4 Then
					SQL3 = "UPDATE pengadaan_pks SET nomor5='" & nomorxx & "' WHERE ID=" & pksid
				Else If a = 5 Then
					SQL3 = "UPDATE pengadaan_pks SET nomor6='" & nomorxx & "' WHERE ID=" & pksid
				End If				
				
				myCmd.CommandText = SQL3
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				myConn.Close()
				
				response.write(nomorxx)
			
			End If

        End If

    Catch ex As Exception
        'lbl1.Text="Error while connecting Server."
		'Response.Write (ex.Message)

    Finally
		'myReader.Close()
        'myConn.Close() 'Whether there is error or not. Close the connection.

    End Try
%>
