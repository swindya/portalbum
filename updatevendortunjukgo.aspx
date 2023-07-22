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
	Dim nomortunx as String = ""
	Dim nomortun2x as String = ""
	Dim nomortun3x as String = ""
	Dim nomortun4x as String = ""
	Dim nomortun5x as String = ""
	Dim nomortun6x as String = ""
	Dim vendor0 as String = ""
	Dim vn0 as Integer = 0
	
	Dim mohonid as Integer = 0
	Dim tunid as Integer = 0
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



	If NOT Request.QueryString("vendor") is Nothing Then
		 vendor0 = Request.QueryString("vendor")
	End If
	
	If NOT Request.QueryString("mohonid") is Nothing Then
		 mohonid = Val(Request.QueryString("mohonid"))
	End If
	
	If NOT Request.QueryString("tunid") is Nothing Then
		 tunid = Val(Request.QueryString("tunid"))
	End If
	
	If NOT Request.QueryString("vno") is Nothing Then
		 vn0 = Val(Request.QueryString("vno"))
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
			
'==============================================================================
'  data
'==============================================================================
			
			If vn0 = 1 Then
				SQL0 = "UPDATE pengadaan_penunjukan SET vendor = '" & vendor0 & "'"
			Else If vn0 = 2 Then
				SQL0 = "UPDATE pengadaan_penunjukan SET vendor2 = '" & vendor0 & "'"
			Else If vn0 = 3 Then
				SQL0 = "UPDATE pengadaan_penunjukan SET vendor3 = '" & vendor0 & "'"
			Else If vn0 = 4 Then
				SQL0 = "UPDATE pengadaan_penunjukan SET vendor4 = '" & vendor0 & "'"
			Else If vn0 = 5 Then
				SQL0 = "UPDATE pengadaan_penunjukan SET vendor5 = '" & vendor0 & "'"
			Else If vn0 = 6 Then
				SQL0 = "UPDATE pengadaan_penunjukan SET vendor6 = '" & vendor0 & "'"
			Else If vn0 = 12 Then
				SQL0 = "UPDATE pengadaan_penunjukan SET vendor2 = ''"
			Else If vn0 = 13 Then
				SQL0 = "UPDATE pengadaan_penunjukan SET vendor3 = ''"
			Else If vn0 = 14 Then
				SQL0 = "UPDATE pengadaan_penunjukan SET vendor4 = ''"
			Else If vn0 = 15 Then
				SQL0 = "UPDATE pengadaan_penunjukan SET vendor5 = ''"
			Else If vn0 = 16 Then
				SQL0 = "UPDATE pengadaan_penunjukan SET vendor6 = ''"
			End If
			
			SQL0 = SQL0 & " WHERE pengadaanID=" & mohonid

			myCmd.CommandText = SQL0
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			myConn.Close()

            If myReader.HasRows Then
				response.write(vn0)
			End If
			
'=====================================================================================================================

        End If

    Catch ex As Exception
        'lbl1.Text="Error while connecting Server."
		'Response.Write (ex.Message)

    Finally
		'myReader.Close()
        'myConn.Close() 'Whether there is error or not. Close the connection.

    End Try
%>
