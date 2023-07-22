<%@ Page Language="VB" %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.OleDb" %> 
<%@ Import Namespace="System.Configuration" %>
<%@ Import Namespace="System.Web.Configuration" %>
 
<script runat="server">

</script>

<%
    'Inherits System.Windows.Forms.Form
    'Create ADO.NET objects.
    Dim myConn As SqlConnection
    Dim myCmd As SqlCommand
	Dim myCmdi As SqlCommand
    Dim myReader As SqlDataReader
    'results As String
	Dim a, aaa, jmla, jmldata as integer
	Dim npp0 as string = ""
	Dim nama0 as string = ""
	Dim posisi0 as String = ""
	Dim jenisoutlet0 as String = ""
	Dim outlet0 as String = ""
	Dim nosk0 as String = ""
	Dim tgl0 as String = ""
	Dim ket0 as String = ""
	Dim namakc as String = ""
	Dim kcid as Integer = 0
	Dim namakcp(200) as String

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
	Dim optionstr as String = ""
	Dim klpid as String = ""
	Dim namaklp as String = ""
	Dim klplist as String = ""

	Dim SQL0 As String
	Dim SQL1 As String
	Dim SQL2 As String
	Dim SQLjo As String
	
	Dim statuskc as Integer = 0
	Dim statuskcp as Integer = 0

	If NOT Request.QueryString("divid") is Nothing Then
		kcx = Request.QueryString("divid")
	End If

	aaa = 1
    Try
        If (aaa=1)
%>

<!--#include file="koneksi.aspx"-->

<%
			myCmd = myConn.CreateCommand

'** Cek Kelompok
			SQLjo = "SELECT * FROM [klp] WHERE (klp_div=" & kcx & ") ORDER BY klp_nama ASC"
			myCmd.CommandText = SQLjo
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			statuskc = 0
            If myReader.HasRows Then
				klplist = klplist + "<option value=''>" + "</option>"
                While myReader.Read()
					statuskc = 1
					klpid = myReader("klp_id").ToString()
					namaklp = myReader("klp_nama").ToString()
					klplist = klplist + "<option value='" + klpid + "'>" + namaklp + "</option>"
				End While
				response.write(klplist)
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

