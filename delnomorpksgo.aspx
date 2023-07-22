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
	Dim pksid as Integer = 0

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
	Dim spkid as Integer = 0
	Dim nosanggahx as String = ""
	Dim nosanggah2x as String = ""
	Dim nosanggah3x as String = ""
	Dim nosanggah4x as String = ""
	Dim nosanggah5x as String = ""
	Dim nosanggah6x as String = ""
	Dim nomor0 as String = ""
	Dim nomorke as Integer = 0
	
	Dim mohonid as Integer = 0
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
	
	If NOT Request.QueryString("nomor") is Nothing Then
		 nomor0 = Request.QueryString("nomor")
	End If
	
	If NOT Request.QueryString("nomorke") is Nothing Then
		 nomorke = Val(Request.QueryString("nomorke"))
	End If
	
	
'***************************************************************************************************************	

	aaa = 1
    Try
        If (aaa=1)
			'myConn = New SqlConnection("Data Source=DESKTOP-S7N8DF9;Initial Catalog=KCT;User ID=sa;Password=sa;")
%>
<!--#include file="koneksi.aspx"-->
<%				
			
			myCmd = myConn.CreateCommand
			

'=====================================================================================================================
' Update nomor pada pengadaan PKS
'=====================================================================================================================
 
 			If nomorke = 2 Then
				SQL1 = "UPDATE pengadaan_pks SET nomor2='' WHERE ID=" & pksid
			Else If nomorke = 3 Then
				SQL1 = "UPDATE pengadaan_pks SET nomor3='' WHERE ID=" & pksid
			Else If nomorke = 4 Then
				SQL1 = "UPDATE pengadaan_pks SET nomor4='' WHERE ID=" & pksid
			Else If nomorke = 5 Then
				SQL1 = "UPDATE pengadaan_pks SET nomor5='' WHERE ID=" & pksid
			Else If nomorke = 6 Then
				SQL1 = "UPDATE pengadaan_pks SET nomor6='' WHERE ID=" & pksid
			End If
			
			myCmd.CommandText = SQL1
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			myConn.Close()
			
			SQL2 = "DELETE FROM surat_pks WHERE nomor='" & nomor0 & "'"
			myCmd.CommandText = SQL2
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			myConn.Close()			

			response.write(SQL2)

        End If

    Catch ex As Exception
        'lbl1.Text="Error while connecting Server."
		'Response.Write (ex.Message)

    Finally
		'myReader.Close()
        'myConn.Close() 'Whether there is error or not. Close the connection.

    End Try
%>
