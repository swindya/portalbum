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
	Dim b as Single = 0
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
	Dim nomornkpx as String = ""
	Dim jenissuratx as String = ""
	
	Dim mohonid as Integer = 0
	Dim nkpid as Integer = 0
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



	If NOT Request.QueryString("perihal") is Nothing Then
		 perihal0 = Request.QueryString("perihal")
	End If
	
	If NOT Request.QueryString("mohonid") is Nothing Then
		 mohonid = Val(Request.QueryString("mohonid"))
	End If
	
	If NOT Request.QueryString("nkpid") is Nothing Then
		 nkpid = Val(Request.QueryString("nkpid"))
	End If
	
	If NOT Request.QueryString("tahun") is Nothing Then
		 tahun = Val(Request.QueryString("tahun"))
	End If
	
	If NOT Request.QueryString("jenissurat") is Nothing Then
		 suratmemo = Request.QueryString("jenissurat")
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
' Cari apakah data NKP sudah punya nomor. Jika belum masukkan data
'==============================================================================
			SQL3 = "SELECT * FROM pengadaan_nkp WHERE ID=" & nkpid
'response.write(SQL3 & "<br>")
			myCmd.CommandText = SQL3
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			a = 0
			If myReader.HasRows Then
				While myReader.Read()
					nomornkpx = myReader("nomor").ToString()
					jenissuratx = myReader("jenissurat").ToString()
				End While
			End If
			myConn.Close()			
			If nomornkpx.length() > 2 Then
				a = 1
			End If
'response.write(a)
'==============================================================================
' Jika nomor BELUM ADA (a=0) -> BIKIN nomor
'==============================================================================
			suratmemo = suratmemo.ToLower()
			If a = 0 Then
				maxu = 0
	'** Cek max value of nomor urut

				If suratmemo="memo" Then
					SQLjo = "SELECT MAX(nourut) AS maxnourut FROM memo WHERE (rubrik LIKE '%" & divstr & "%' AND jenis=1 AND tahun=" & tahun & ")"
				Else If suratmemo="notin" Then
					SQLjo = "SELECT MAX(nourut) AS maxnourut FROM notin WHERE (rubrik LIKE '%" & divstr & "%' AND tahun=" & tahun & ")"
					SQLjo = "SELECT MAX(nourut) AS maxnourut FROM notin WHERE (tahun=" & tahun & ")"
				End If
'response.write(SQLjo)
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
				Dim nomorq as String = "OPR/4.3/" & maxu
				perihal0 = "NKP: " & perihal0
				If suratmemo="memo" Then
					SQL0 = "INSERT INTO memo (nourut, tglmemo, tglinput, jenis, rubrik, tahun, klp, createduser," & _
							"perihal, nomor, jam, divisi, createddate) VALUES(" & maxu & ",'" & tgl0 & "','" & _
							tglinputq & "',1,'OPR/4.3/'," & tahun & "," & klpq & "," & idq & ",'" & _
							perihal0.ToUpper() & "','" & nomorq & "','" & jamsaiki & "'," & divq & ",'" & _
							tglinputq & "')"
					SQL1 = "DELETE FROM notin WHERE nomor='" & nomorq & "'"
				Else If suratmemo="notin" Then
					SQL0 = "INSERT INTO notin (nourut, tglnotin, tglinput, jenis, rubrik, tahun, klp, createduser," & _
							"perihal, nomor, jam, divisi, createddate) VALUES(" & maxu & ",'" & tgl0 & "','" & _
							tglinputq & "',1,'OPR/4.3/'," & tahun & "," & klpq & "," & idq & ",'" & _
							perihal0.ToUpper() & "','" & nomorq & "','" & jamsaiki & "'," & divq & ",'" & _
							tglinputq & "')"
					SQL1 = "DELETE FROM memo WHERE nomor='" & nomorq & "'"
				End If
'response.write(SQL0)			
				myCmd.CommandText = SQL0
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				myConn.Close()
'=====================================================================================================================
				If suratmemo="memo" Then
					SQL2 = "SELECT * FROM [memo] WHERE (nourut=" & maxu & " AND jenis=1 AND tahun=" & tahun & " AND rubrik='OPR/4.3/')"
				Else If suratmemo="notin" Then
					SQL2 = "SELECT * FROM [notin] WHERE (nourut=" & maxu & " AND tahun=" & tahun & ")"
				End If
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
' Update nomor pada pengadaannkp
'=====================================================================================================================
				SQL3 = "UPDATE pengadaan_nkp SET nomor='" & nomorq & "', jenissurat='" & suratmemo & "' WHERE ID=" & nkpid
				myCmd.CommandText = SQL3
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				myConn.Close()
				
				response.write(nomorq)

'==============================================================================
' Jika nomor SUDAH ADA (a=0) -> GANTI nomor
'==============================================================================

			Else If a > 0 Then
'Periksa apakah jenis surat sama/beda
'Jika jenis surat beda -> baca jenis surat di nkp
'response.write(jenissuratx & "--" & suratmemo)
				If jenissuratx <> suratmemo Then
'Jika beda -> hapus memo terkait -> bikin nomor baru
					If suratmemo="memo" Then
						SQL1 = "DELETE FROM notin WHERE nomor='" & nomornkpx & "'"
					Else If suratmemo="notin" Then
						SQL1 = "DELETE FROM memo WHERE nomor='" & nomornkpx & "'"
					End If				
					myCmd.CommandText = SQL1
					myConn.Open()
					myReader = myCmd.ExecuteReader()
					myConn.Close()
					
					If suratmemo="memo" Then
						SQLjo = "SELECT MAX(nourut) AS maxnourut FROM memo WHERE (rubrik LIKE '%" & divstr & "%' AND jenis=1 AND tahun=" & tahun & ")"
					Else If suratmemo="notin" Then
						SQLjo = "SELECT MAX(nourut) AS maxnourut FROM notin WHERE (rubrik LIKE '%" & divstr & "%' AND tahun=" & tahun & ")"
						SQLjo = "SELECT MAX(nourut) AS maxnourut FROM notin WHERE (tahun=" & tahun & ")"
					End If

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
'=============================================================================
					Dim nomorq as String = "OPR/4.3/" & maxu
					perihal0 = "NKP: " & perihal0
					If suratmemo="memo" Then
						SQL0 = "INSERT INTO memo (nourut, tglmemo, tglinput, jenis, rubrik, tahun, klp, createduser," & _
								"perihal, nomor, jam, divisi, createddate) VALUES(" & maxu & ",'" & tgl0 & "','" & _
								tglinputq & "',1,'OPR/4.3/'," & tahun & "," & klpq & "," & idq & ",'" & _
								perihal0.ToUpper() & "','" & nomorq & "','" & jamsaiki & "'," & divq & ",'" & _
								tglinputq & "')"
					Else If suratmemo="notin" Then
						SQL0 = "INSERT INTO notin (nourut, tglnotin, tglinput, jenis, rubrik, tahun, klp, createduser," & _
								"perihal, nomor, jam, divisi, createddate) VALUES(" & maxu & ",'" & tgl0 & "','" & _
								tglinputq & "',1,'OPR/4.3/'," & tahun & "," & klpq & "," & idq & ",'" & _
								perihal0.ToUpper() & "','" & nomorq & "','" & jamsaiki & "'," & divq & ",'" & _
								tglinputq & "')"
					End If
					myCmd.CommandText = SQL0
					myConn.Open()
					myReader = myCmd.ExecuteReader()
					myConn.Close()
					
'=====================================================================================================================
' Update nomor pada pengadaannkp
'=====================================================================================================================
					SQL3 = "UPDATE pengadaan_nkp SET nomor='" & nomorq & "', jenissurat='" & suratmemo & "' WHERE ID=" & nkpid
					myCmd.CommandText = SQL3
					myConn.Open()
					myReader = myCmd.ExecuteReader()
					myConn.Close()
					
					response.write(nomorq)
'Jika sama -> nothing done		
				End If
			End If

        End If

    Catch ex As Exception
        'lbl1.Text="Error while connecting Server."
		Response.Write (ex.Message)

    Finally
		'myReader.Close()
        'myConn.Close() 'Whether there is error or not. Close the connection.

    End Try
%>
