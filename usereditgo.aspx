<%@ Page Language="VB" %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.OleDb" %> 
<%@ Import Namespace="System.Configuration" %>
<%@ Import Namespace="System.Web.Configuration" %>

<html>
<head>
<style>
.verticalhorizontal {
  margin: 0;
  position: absolute;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
}
</style>
</head>
<body>
<%
    'Inherits System.Windows.Forms.Form
    'Create ADO.NET objects.
    Dim myConn As SqlConnection
    Dim myCmd As SqlCommand
	Dim myCmdi As SqlCommand
    Dim myReader As SqlDataReader
    'results As String
	Dim a, aaa, jmla, jmldata as integer
	Dim nama1 as string = ""
	Dim npp1 as string = ""
	Dim div1 as String = ""
	Dim kel1 as String = ""
	Dim unit1 as String = "0"
	Dim posisi1 as String = ""
	Dim level1 as String = ""
	Dim csurat1 as String = ""
	Dim csekre1 as String = ""
	Dim cada1 as String = ""
	Dim caset1 as String = ""
	Dim pwd11 as String = ""
	Dim pwd21 as String = ""
	Dim id1 as String = ""

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
	Dim SQLjo As String
	
	Dim statusexist as Integer = 0
	
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
	
	Dim divstrc as String = ""
	Dim kelstrc as String = ""
	Dim unitstrc as String = ""
	Dim posisistrc as String = ""

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

'Waktu -> prefix namafile
	'Dim s As String = DateTime.Now.ToString("yyyyMMdd")
	's = s & Hour(Now) & Minute(Now) & Second(Now)
	'namatemp = nama0.ToLower()
	'namatemp = namatemp.Replace(" ", "")
	'prenamafile = namatemp.Substring(0, 6) & s

	If NOT Request.Form("editid") is nothing Then
		id1 = Request.Form("editid")
	End If
	
	If NOT Request.Form("editnama") is nothing Then
		nama1 = Request.Form("editnama")
	End If
	
	If NOT Request.Form("editnpp") is nothing Then
		npp1 = Request.Form("editnpp")
	End If
	
	If NOT Request.Form("editdivisi") is nothing Then
		div1 = Request.Form("editdivisi")
	End If

	If NOT Request.Form("editkelompok") is nothing Then
		kel1 = Request.Form("editkelompok")
	End If
	
	If NOT Request.Form("editunit") is nothing Then
		unit1 = Request.Form("editunit")
	End If
	
	If NOT Request.Form("editposisi") is nothing Then
		posisi1 = Request.Form("editposisi")
	End If
	
	If NOT Request.Form("editlevel") is nothing Then
		level1 = Request.Form("editlevel")
	End If

	If NOT Request.Form("editcsurat") is nothing Then
		csurat1 = Request.Form("editcsurat")
	End If

	If NOT Request.Form("editcsekre") is nothing Then
		csekre1 = Request.Form("editcsekre")
	End If
	
	If NOT Request.Form("editcada") is nothing Then
		cada1 = Request.Form("editcada")
	End If

	If NOT Request.Form("editcaset") is nothing Then
		caset1 = Request.Form("editcaset")
	End If
	
	If NOT Request.Form("editpwd") is nothing Then
		pwd11 = Request.Form("editpwd")
	End If
	
	If NOT Request.Form("editpwd2") is nothing Then
		pwd21 = Request.Form("editpwd2")
	End If
	
'	If NOT Request.Form("barutgl") is nothing Then
'		tgl0 = Request.Form("barutgl")
'		if tgl0.length() > 5 Then
'			Dim tglarr(3) as String 
'			tglarr = tgl0.Split("-")
'			tgl0 = tglarr(2) & "-" & tglarr(1) & "-" & tglarr(0)
'			tahun = Cint(tglarr(2))
'			bulan = Cint(tglarr(1))
'			tglhari = Cint(tglarr(0))
'		End If
'	End If
	
'Waktu -> prefix namafile
	Dim s As String = DateTime.Now.ToString("yyyyMMdd")
	Dim tglinputq As String = DateTime.Now.ToString("yyyy-MM-dd")
	Dim jamsaiki As String = DateTime.Now.ToString("HH:mm:ss")
	Dim tahunstr as String = DateTime.Now.ToString("yyyy")
	Dim tahunq as Integer = Val(tahunstr)
	s = s & Hour(Now) & Minute(Now) & Second(Now)
	prenamafile = s

	'folder1 = "./DOC/surat/" & tahun & "/" & bulan & "/" & tglhari & "/"
	
	'Dim folderPath As String = Server.MapPath(folder1)
     'Check whether Directory (Folder) exists.

	'If(Not System.IO.Directory.Exists(folderPath)) Then
	'	Directory.CreateDirectory(folderPath)
	'End If

    'Save the File to the Directory (Folder).
	'Dim exten As String = ""
	'Dim postedFile As HttpPostedFile = Request.Files("uploadImage0")
	'Dim namafolderx as String = "/DOC/surat/" & tahun & "/" & bulan & "/"
	

'response.write(namafilexx)
	
'***************************************************************************************************************	

	aaa = 1
    Try
        If (aaa=1)
%>
<!--#include file="koneksi.aspx"-->
<%				
			
			myCmd = myConn.CreateCommand

'** Cek max value of nomor urut
			SQLjo = "SELECT * FROM [user] WHERE (idemployee='" & npp1 & "')"
			a=0
			myCmd.CommandText = SQLjo
			myConn.Open()
			myReader = myCmd.ExecuteReader()
            If myReader.HasRows Then
				While myReader.Read()
					a = a + 1
				End While
			End If
			myConn.Close()
			if a > 0 Then
				statusexist=1
			End If		
'==========================================================================
' Cari Daftar Divisi
'==========================================================================
			SQL0 = "SELECT * FROM [divisi] WHERE div_id=" & div1
			myCmd.CommandText = SQL0
			myConn.Open()
			myReader = myCmd.ExecuteReader()
            If myReader.HasRows Then
                While myReader.Read()
					divstrc = myReader("div_nama").ToString()
				End While
			End If
			myConn.Close()
'==========================================================================
' Cari Daftar Kelompok
'==========================================================================
			SQL0 = "SELECT * FROM [klp] WHERE klp_id=" & kel1
			myCmd.CommandText = SQL0
			myConn.Open()
			myReader = myCmd.ExecuteReader()
            If myReader.HasRows Then
                While myReader.Read()
					kelstrc = myReader("klp_nama").ToString()
				End While
			End If
			myConn.Close()
'response.write(kelstrc & "<br>")
'==========================================================================
' Cari Daftar Unit
'==========================================================================
			unitstrc = ""
			If Not unit1 is Nothing AND unit1.length() > 0 Then
				SQL0 = "SELECT * FROM [unit] WHERE u_id=" & unit1
				myCmd.CommandText = SQL0
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				b = 0
				If myReader.HasRows Then
					While myReader.Read()
						b = b + 1
						unitstrc = myReader("u_nama").ToString()
					End While
				End If
				myConn.Close()
			End If
'==========================================================================
' Cari Daftar Posisi
'==========================================================================
			SQL0 = "SELECT * FROM [posisi] WHERE po_id=" & posisi1
			myCmd.CommandText = SQL0
			myConn.Open()
			myReader = myCmd.ExecuteReader()
            If myReader.HasRows Then
                While myReader.Read()
					posisistrc = myReader("posisi").ToString()
				End While
			End If
			myConn.Close()
'==========================================================================
'response.write("------------------------------------------" & SQLjo)

			If statusexist=1 Then
				If pwd11="11111111" Then
					SQL0 = "UPDATE [user] SET nama='" & nama1 & "', idemployee='" & npp1 & "', div=" & div1 & ", klp=" & kel1 & _
							", unit=" & unit1 & ", posisi=" & posisi1 & ", menusurat=" & csurat1 & ", menusekretaris=" & csekre1 & _
							", menupengadaan=" & cada1 & ", menuaset=" & caset1 & ", levelakses=" & level1 & _
							", posisistr='" & posisistrc & "', divisistr='" & divstrc & "', kelompokstr='" & kelstrc & "', unitstr='" & _
							unitstrc & "' WHERE ID=" & id1
				Else
					SQL0 = "UPDATE [user] SET nama='" & nama1 & "', idemployee='" & npp1 & "', div=" & div1 & ", klp=" & kel1 & _
							", unit=" & unit1 & ", posisi=" & posisi1 & ", menusurat=" & csurat1 & ", menusekretaris=" & csekre1 & _
							", menupengadaan=" & cada1 & ", menuaset=" & caset1 & ", levelakses=" & level1 & _
							", paswd=HASHBYTES('SHA2_256','" & pwd11 & "')" & _
							", posisistr='" & posisistrc & "', divisistr='" & divstrc & "', kelompokstr='" & kelstrc & "', unitstr='" & _
							unitstrc & "' WHERE ID=" & id1				
				End If
'response.write(SQL0 & "<br>")
				myCmd.CommandText = SQL0
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				myConn.Close()
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
<img src="./images/spinnercolor.gif" alt="spinner" class="verticalhorizontal" height="200" width="200">
<meta http-equiv="refresh" content="0;url=user.aspx" />
</body>
</html>