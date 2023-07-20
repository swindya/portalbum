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
	Dim a, b, aaa, jmla, jmldata as integer

	'Dim pageno as Integer = 1
	Dim pageno as String
	Dim jmlpage as Integer = 1
	Dim pagerows as Integer = 20
	Dim rowsperpage as Integer = 20
	Dim hal as Integer = 1
	Dim c as Integer = 0
	Dim idx as Integer = 0
	Dim itop as Integer = 1
	Dim ii as Integer = 0
	Dim k as Integer
	Dim userid as Integer = 0
	Dim namauserx as String = ""

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
	Dim levelaksesq as String = ""
	Dim mohonid as Integer = 0
	Dim tglku as String
	Dim jenisver as String
	Dim status as String
	Dim ket as String
	Dim dataexist as Integer = 0
	Dim tglverx as String
	Dim statusup as Integer = 0
	Dim nourut as Integer = 0

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


	If NOT Request.Form("verpadaid") is Nothing Then
		mohonid = Cint(Request.Form("verpadaid"))
	End If
	
	If NOT Request.Form("vertgl") is Nothing Then
		tglku = Request.Form("vertgl")
		Dim tglstr() = tglku.Split("-")
		tglverx = tglstr(2) & "-" & tglstr(1) & "-" & tglstr(0)
	End If
	
	If NOT Request.Form("verjenis") is Nothing Then
		jenisver = Request.Form("verjenis")
		jenisver = jenisver.ToUpper()
	End If
	
	If NOT Request.Form("verstatus") is Nothing Then
		status = Request.Form("verstatus")
		status = status.ToUpper()
	End If
	
	If NOT Request.Form("verket") is Nothing Then
		ket = Request.Form("verket")
		'ket = status.ToUpper()
	End If
	
	levelaksesq = Session("levelakses")
	'posisi20 = Left(posisix, 20)
	
'***************************************************************************************************************


	aaa = 1
    Try
        If aaa=1 Then
%>
<!--#include file="koneksi.aspx"-->
<%				
			
			myCmd = myConn.CreateCommand
			nourut = 0
'** Cek apakah data sudah ada
			SQLjo = "SELECT * FROM [pengadaanusulverifikasi] WHERE (pengadaanID=" & mohonid & " AND jenis='" & jenisver & "') ORDER BY nourut ASC"
			myCmd.CommandText = SQLjo
			myConn.Open()
			myReader = myCmd.ExecuteReader()
            If myReader.HasRows Then
                While myReader.Read()
					nourut = nourut + 1
					'idx = myReader("nourut").ToString()
				End While
			End If
			myConn.Close()
			nourut = nourut + 1
			
'==================================================================================================================
' Masukkan ke database
'==================================================================================================================
			SQL0 = "INSERT INTO pengadaanusulverifikasi (pengadaanID, tgl, jenis, status, nourut, keterangan, createduser, createddatetime) VALUES(" & _
					mohonid & ",'" & tglverx & "','" & jenisver & "','" & status & "'," & nourut & ",'" & ket & "'," & idq & ",GETDATE())"
					'kcx.ToUpper() & "',1)"
'response.write(SQL0)			
			myCmd.CommandText = SQL0
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			myConn.Close()
			
        End If

    Catch ex As Exception
        'lbl1.Text="Error while connecting Server."
		'Response.Write (ex.Message)

    Finally
		'myReader.Close()
        'myConn.Close() 'Whether there is error or not. Close the connection.

    End Try
	
	Session("username") = user0
	Session("userid") = idq
	Session("namauser") = namauserq
	'Session("idemployee") = nppq
	Session("levelakses") = levelaksesq
	Session("klp") = klpq
	Session("unit") = unitq
	Session("posisi") = posisiq
	Session("posisistr") = posisix
	Session("divisi") = divq

			
%>
<img src="./images/spinnercolor.gif" alt="spinner" class="verticalhorizontal" height="200" width="200">
<meta http-equiv="refresh" content="0;url=prosesusulan.aspx?padaid=<% response.write(mohonid) %>" />
</body>
</html>