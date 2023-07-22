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
	Dim perihal0 as string = ""
	Dim nama0 as string = ""
	Dim posisi0 as String = ""
	Dim jenisund0 as String = ""
	Dim ada as String = ""
	Dim tembusan0 as String = ""
	Dim pengirim0 as String = ""
	Dim tgl0 as String = ""
	Dim jenisnotin0 as String = ""
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
	Dim notinq As String = ""
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
	Dim jenisanggaran0 as Integer = 0
	Dim notinidx as Long = 0
	Dim divstr as String = ""
	Dim idku as Integer = 0

	user0 = Session("username")
	idq = Session("userid")
	namauserq = Session("namauser")
	jabatanq = Session("jabatan")
	klpq = Session("klp")
	unitq = Session("unit")
	posisiq = Session("posisi")
	posisix = Session("posisistr")
	notinq = Session("notin")
	divq = Session("divisi")

'Waktu -> prefix namafile
	'Dim s As String = DateTime.Now.ToString("yyyyMMdd")
	's = s & Hour(Now) & Minute(Now) & Second(Now)
	'namatemp = nama0.ToLower()
	'namatemp = namatemp.Replace(" ", "")
	'prenamafile = namatemp.Substring(0, 6) & s


	If NOT Request.Form("barujenis") is nothing Then
		jenis0 = Request.Form("barujenis")
	End If
	
	If NOT Request.Form("baruperihal") is nothing Then
		perihal0 = Request.Form("baruperihal")
	End If
	
	If NOT Request.Form("barupengirim") is nothing Then
		pengirim0 = Request.Form("barupengirim")
	End If

	If NOT Request.Form("barutujuan") is nothing Then
		tujuan0 = Request.Form("barutujuan")
	End If
	
	If NOT Request.Form("barutembusan") is nothing Then
		tembusan0 = Request.Form("barutembusan")
	End If

	If NOT Request.Form("baruju") is nothing Then
		jenisund0 = Request.Form("baruju")
	End If
	
	If NOT Request.Form("baruada") is nothing Then
		ada = Request.Form("baruada")
	End If
	
	If NOT Request.Form("barutgl") is nothing Then
		tgl0 = Request.Form("barutgl")
		if tgl0.length() > 5 Then
			Dim tglarr(3) as String 
			tglarr = tgl0.Split("-")
			tgl0 = tglarr(2) & "-" & tglarr(1) & "-" & tglarr(0)
			tahun = Cint(tglarr(2))
			bulan = Cint(tglarr(1))
			tglhari = Cint(tglarr(0))
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

	'folder1 = "~/portalbum/DOC/notaintern/" & tahun & "/" & bulan & "/" & tglhari & "/"
	folder1 = "./DOC/notaintern/" & tahun & "/" & bulan & "/" & tglhari & "/"
	folder2 = "./DOC/pengadaan/" & tahun & "/" & bulan & "/" & tglhari & "/"
	
	Dim folderPath As String = Server.MapPath(folder1)
     'Check whether Directory (Folder) exists.

	If(Not System.IO.Directory.Exists(folderPath)) Then
		'System.IO.Directory.CreateDirectory(folderPath)
		Directory.CreateDirectory(folderPath)
	End If

    'Save the File to the Directory (Folder).
	Dim exten As String = ""
	Dim postedFile As HttpPostedFile = Request.Files("uploadImage0")
	Dim namafolderx as String = "/DOC/notaintern/" & tahun & "/" & bulan & "/" & tglhari & "/"
	Dim namafolderadax as String = "/DOC/pengadaan/" & tahun & "/" & bulan & "/" & tglhari & "/"
	
    'Check if File is available.
	Dim filearr(10) as String
	Dim namafilexx As String = ""
	If postedFile IsNot Nothing And postedFile.ContentLength > 0 Then
        'Save the File.
        'Dim filePath1 As String = Server.MapPath("~/DOC/portalbum/") & Path.GetFileName(postedFile.FileName)
		'Dim filePath1 As String = Server.MapPath(folder1) + postedFile.FileName
		Dim namafilex = postedFile.FileName
		filearr = namafilex.Split(".")
		Dim filearrno as Integer = filearr.length
		Dim aa = filearrno - 2
		namafilexx = ""
		For a=0 to aa
			namafilexx = namafilexx & filearr(a)
'response.write(namafilexx & "<br>")
		Next a
		exten = System.IO.Path.GetExtension(postedFile.FileName)
		namafilexx = prenamafile & namafilexx & exten
		namafile0 = Server.MapPath(folder1) + namafilexx
        postedFile.SaveAs(namafile0)
		If Val(ada) = 1 Then
			namafile0 = Server.MapPath(folder2) + namafilexx
			postedFile.SaveAs(namafile0)		
		End If
    End If
	

'response.write(namafilexx)

	Dim divarr(10) as String
	divstr = ""
	If pengirim0.length > 1 Then
		divarr = pengirim0.Split("/")
		divstr = divarr(0)
	End If
'***************************************************************************************************************	

	aaa = 1
    Try
        If aaa=1 Then
			'myConn = New SqlConnection("Data Source=DESKTOP-S7N8DF9;Initial Catalog=KCT;User ID=sa;Password=sa;")
%>
<!--#include file="koneksi.aspx"-->
<%				
			
			myCmd = myConn.CreateCommand

'** Cek max value of nomor urut
			'SQLjo = "SELECT MAX(nourut) AS maxnourut FROM [notin] WHERE (rubrik='" & pengirim0 & "' AND tahun=" & tahun & ")"
			'SQLjo = "SELECT MAX(nourut) AS maxnourut FROM [notin] WHERE (rubrik LIKE '%" & divstr & "%' AND jenis=" & jenis0 & _
			'		" AND tahun=" & tahun & ")"
			SQLjo = "SELECT MAX(nourut) AS maxnourut FROM [notin] WHERE (rubrik LIKE '%" & divstr & "%' AND tahun=" & tahun & ")"
'response.write(SQLjo)
			maxu = 0
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
			
'==================================================================================================================
' Masukkan ke Nota Intern
'==================================================================================================================
			Dim nomorq as String = pengirim0 & maxu
			'SQL0 = "INSERT INTO notin (nourut, tglnotin, tglinput, jenis, rubrik, tahun, klp, createduser, tujuan," & _
			'		"perihal, cc, nomor, jam, divisi, createddate, namafile, namafolder, undangan, statuspengadaan) VALUES(" & _
			'		maxu & ",'" & tgl0 & "','" & tglinputq & "'," & jenis0 & ",'" & pengirim0 & "'," & tahun & "," & klpq & "," & _
			'		idq & ",'" & tujuan0 & "','" & perihal0 & "','" & tembusan0 & "','" & nomorq & "','" & jamsaiki & "'," & divq & ",'" & 
			'		tglinputq & "','" & namafilexx & "','" & namafolderx & "'," & jenisund0 & "," & ada & ")"
			SQL0 = "INSERT INTO notin (nourut, tglnotin, tglinput, rubrik, tahun, klp, createduser, tujuan," & _
					"perihal, cc, nomor, jam, divisi, createddate, namafile, namafolder, undangan, statuspengadaan) VALUES(" & _
					maxu & ",'" & tgl0 & "','" & tglinputq & "','" & pengirim0.ToUpper() & "'," & tahun & "," & klpq & "," & _
					idq & ",'" & tujuan0.ToUpper() & "','" & perihal0.ToUpper() & "','" & tembusan0.ToUpper() & "','" & nomorq & "','" & jamsaiki & "'," & divq & ",'" & 
					tglinputq & "','" & namafilexx & "','" & namafolderx & "'," & jenisund0 & "," & ada & ")"
					'kcx.ToUpper() & "',1)"
'response.write(SQL0 & "<br>")			
			myCmd.CommandText = SQL0
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			myConn.Close()

'==================================================================================================================
' Cari notinID -> Masukkan ke pengadaan
'==================================================================================================================
			SQL1 = "SELECT ID AS maxid FROM [notin] WHERE (nomor='" & nomorq & "')"
			myCmd.CommandText = SQL1
			myConn.Open()
			myReader = myCmd.ExecuteReader()
            If myReader.HasRows Then
                While myReader.Read()
					notinidx = Val(myReader("maxid").ToString())
				End While
			End If
			myConn.Close()

'==================================================================================================================
' Masukkan ke Pengadaan
'==================================================================================================================
			If Val(ada) = 1 Then
				'SQL0 = "INSERT INTO pengadaan (notinID, tglmohon, nama, klp, divisi, jenispengadaan, jenisanggaran, tahun, " & _
				'		"statuspass, nostep, step, createduser, createddate, statusrevisi, namafolder, namafilesurat) VALUES(" & notinidx & _
				'		",'" & tgl0 & "','" & perihal0.ToUpper() & "'," & klpq & "," &	divq & ",'" & jenis0 & "','" & jenisanggaran0 & _
				'		"'," & tahun & ",0,1,'IJIN'," & idq & ",GETDATE(),0,'" & namafolderx & "','" & namafilexx & "')"
				SQL0 = "INSERT INTO pengadaan (notinID, tglmohon, nama, klp, divisi, tahun, " & _
						"statuspass, nostep, step, createduser, createddate, statusrevisi, namafolder, namafilesurat) VALUES(" & notinidx & _
						",'" & tgl0 & "','" & perihal0.ToUpper() & "'," & klpq & "," &	divq & _
						"," & tahun & ",0,1,'IJIN'," & idq & ",GETDATE(),0,'" & namafolderadax & "','" & namafilexx & "')"	
'response.write(SQL0 & "<br>")
				myCmd.CommandText = SQL0
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				myConn.Close()
			End If
'=========================================
			SQL2 = "SELECT * FROM [notin] WHERE (nourut=" & maxu & " AND jenis=" & jenis0 & " AND rubrik='" & pengirim0 & "')"
			myCmd.CommandText = SQL2
			myConn.Open()
			myReader = myCmd.ExecuteReader()
            If myReader.HasRows Then
                While myReader.Read()
					idku = myReader("ID").ToString()
				End While
			End If
			myConn.Close()			
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
<meta http-equiv="refresh" content="0;url=notainternprint.aspx?id=<% response.write(notinidx) %>" />
</body>
</html>