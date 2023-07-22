<!DOCTYPE html>

<%@ Page Language="VB" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.OleDb" %> 
<%@ Import Namespace="System.Configuration" %>
<%@ Import Namespace="System.Web.Configuration" %>
 
<script runat="server">

</script>

<html>
<head>
    <meta charset="utf-8">
    <!-- Somehow I got an error, so I comment the title, just uncomment to show -->
    <!-- <title>Animated Sidebar Menu | CodingLab</title> -->
    <link rel="stylesheet" href="style.css">
    <script src="https://kit.fontawesome.com/a076d05399.js"></script>
<!-- Font Awesome JS -->
<script src="https://kit.fontawesome.com/b99e675b6e.js"></script>
<!-- jQuery -->
<script src="https://code.jquery.com/jquery-3.4.1.min.js"></script>

    <!--script src="./js/a076d05399.js"></script>
	<script src="./js/b99e675b6e.js"></script>
	<script src="./js/jquery-3.4.1.min.js"></script-->

<link rel="stylesheet" href="css1/styles.css">

<link href="./datepicker4/dcalendar.picker.css" rel="stylesheet" type="text/css">


<style>
.example-element {
	margin-top: 0;
	position: fixed;
	background-color: #0A9DE7;
	width: 230px;
	height: 100%;
}
</style>

<style>
body {
  margin: 0;
  font-family: Arial, Helvetica, sans-serif;
}

.topnav {
  overflow: hidden;
  background-color: #0A9DE7;
  //position: relative;
  margin-top: 0;
}

.topnav a {
  float: right;
  display: block;
  color: #f2f2f2;
  text-align: center;
  padding: 17px 8px;
  margin-top: 5px;
  text-decoration: none;
  font-size: 17px;
}

.topnav a:hover {
  //background-color: #ddd;
  //color: black;
}

.topnav a.active {
  background-color: #4CAF50;
  color: white;
}

.topnav .icon {
  display: none;
}

@media screen and (max-width: 600px) {
  .topnav a:not(:first-child) {display: none;}
  .topnav a.icon {
    float: left;
    display: block;
  }
}

@media screen and (max-width: 600px) {
  .topnav.responsive {position: relative;}
  .topnav.responsive .icon {
    position: absolute;
    right: 0;
    top: 0;
  }
  .topnav.responsive a {
    float: none;
    display: block;
    text-align: left;
  }
}
</style>

<style>
.container {
  position: relative;
  text-align: left;
  color: white;
  font-family: Calibri;
  font-size: 0.9em;
}

.logo {
    left: 20px;
    line-height: 100px;
    margin-top: 500px;
	margin-left: 0px;
    position: absolute;
    text-align: left;
    //top: 50%;
    //width: 100%;
}

.welcome {
    left: 20px;
    line-height: 100px;
    margin-top: -40px;
	margin-left: 0px;
    position: absolute;
    text-align: left;
    //top: 50%;
    //width: 100%;
}
.userin {
    left: 0px;
    line-height: 100px;
	margin-top: -5px;
    margin-left: 70px;
    position: absolute;
    text-align: left;
	color: yellow;
    //top: 50%;
    //width: 100%;
}
.posisi {
    left: 0px;
    line-height: 100px;
	margin-top: 20px;
    margin-left: 70px;
    position: absolute;
    text-align: left;
	color: #D7F5FF;
    //top: 50%;
    //width: 100%;
}
.photo {
    left: 0px;
    line-height: 100px;
	margin-top: 20px;
    margin-left: 20px;
    position: absolute;
    text-align: left;
	color: yellow;
    //top: 50%;
    //width: 100%;
}
</style>

<style>
.button, .button:visited {
background-color: #222;
display: inline-block;
padding: 5px 10px 6px;
color: #fff;
text-decoration: none;
-moz-border-radius: 6px;
-webkit-border-radius: 6px;
-moz-box-shadow: 0 1px 3px rgba(0,0,0,0.6);
-webkit-box-shadow: 0 1px 3px rgba(0,0,0,0.6);
text-shadow: 0 -1px 1px rgba(0,0,0,0.25);
border-bottom: 1px solid rgba(0,0,0,0.25);
position: relative;
cursor: pointer;
font-family: Arial;

}
.button:hover {
background-color: #111;
color: #fff;
} .button:active {
top: 1px;
} 
.small.button, .small.button:visited {
font-size: 0.8em;
padding: 8px 14px 6px;
} 
.button, .button:visited,
.medium.button, .medium.button:visited {
font-size: 1.2em;
font-family: calibri;
font-weight: normal;
line-height: 1;
text-shadow: 0 -1px 1px rgba(0,0,0,0.25); 
} 
.medium.button, .medium.button:visited {
font-size: 1.1em;
padding: 6px 14px 8px;
} 
.large.button, .large.button:visited {
font-size: 1.1em;
padding: 18px 22px 16px;
} 
.pink.button, .magenta.button:visited {
background-color: #e22092;
} 
.pink.button:hover {
background-color: #c81e82;
} 
.green.button, .green.button:visited {
background-color: #00AC04;
} 
.green.button:hover {
background-color: #749a02;
} 
.red.button, .red.button:visited {
background-color: #e62727;
}
.red.button:hover {
background-color: #cf2525;
}
.orange.button, .orange.button:visited {
background-color: #ff5c00;
} 
.orange.button:hover {
background-color: #d45500;
}
.blue.button, .blue.button:visited {
background-color: #2981e4;
}
.blue.button:hover {
background-color: #2575cf;
}
.yellow.button, .yellow.button:visited {
background-color: #ffb515;
}
.yellow.button:hover {
background-color: #fc9200;
} 
</style>
<style> 
.textbox { 
    border: 1px solid #c4c4c4; 
    height: 26px; 
    width: 240px; 
    font-size: 13px; 
    padding: 4px 4px 4px 4px; 
    border-radius: 4px; 
    -moz-border-radius: 4px; 
    -webkit-border-radius: 4px; 
    box-shadow: 0px 0px 4px #d9d9d9; 
    -moz-box-shadow: 0px 0px 4px #d9d9d9; 
    -webkit-box-shadow: 0px 0px 4px #d9d9d9; 
} 
 
.textbox:focus { 
    outline: none; 
    border: 1px solid #7bc1f7; 
    box-shadow: 0px 0px 4px #7bc1f7; 
    -moz-box-shadow: 0px 0px 4px #7bc1f7; 
    -webkit-box-shadow: 0px 0px 4px #7bc1f7; 
</style>
<style>
.textboxnarrow { 
    border: 1px solid #c4c4c4; 
    height: 26px; 
    width: 50px; 
    font-size: 13px; 
    padding: 4px 4px 4px 4px; 
    border-radius: 4px; 
    -moz-border-radius: 4px; 
    -webkit-border-radius: 4px; 
    box-shadow: 0px 0px 4px #d9d9d9; 
    -moz-box-shadow: 0px 0px 4px #d9d9d9; 
    -webkit-box-shadow: 0px 0px 4px #d9d9d9; 
} 
 
.textboxnarrow:focus { 
    outline: none; 
    border: 1px solid #7bc1f7; 
    box-shadow: 0px 0px 4px #7bc1f7; 
    -moz-box-shadow: 0px 0px 4px #7bc1f7; 
    -webkit-box-shadow: 0px 0px 4px #7bc1f7; 
</style>
<style>
.textboxmid { 
    border: 1px solid #c4c4c4; 
    height: 26px; 
    width: 130px; 
    font-size: 13px; 
    padding: 4px 4px 4px 4px; 
    border-radius: 4px; 
    -moz-border-radius: 4px; 
    -webkit-border-radius: 4px; 
    box-shadow: 0px 0px 4px #d9d9d9; 
    -moz-box-shadow: 0px 0px 4px #d9d9d9; 
    -webkit-box-shadow: 0px 0px 4px #d9d9d9; 
} 
 
.textboxmid:focus { 
    outline: none; 
    border: 1px solid #c4c4c4; 
    box-shadow: 0px 0px 4px #7bc1f7; 
    -moz-box-shadow: 0px 0px 4px #7bc1f7; 
    -webkit-box-shadow: 0px 0px 4px #7bc1f7; 
} 
 </style>
 <style>
.textboxwide { 
    border: 1px solid #c4c4c4; 
    height: 26px; 
    width: 250px; 
    font-size: 13px; 
    padding: 4px 4px 4px 4px; 
    border-radius: 4px; 
    -moz-border-radius: 4px; 
    -webkit-border-radius: 4px; 
    box-shadow: 0px 0px 4px #d9d9d9; 
    -moz-box-shadow: 0px 0px 4px #d9d9d9; 
    -webkit-box-shadow: 0px 0px 4px #d9d9d9; 
} 
 
.textboxwide:focus { 
    outline: none; 
    border: 1px solid #7bc1f7; 
    box-shadow: 0px 0px 4px #7bc1f7; 
    -moz-box-shadow: 0px 0px 4px #7bc1f7; 
    -webkit-box-shadow: 0px 0px 4px #7bc1f7; 
} 
 </style>
 <style> 
  .textboxtgl { 
    border: 1px solid #c4c4c4; 
    height: 26px; 
    width: 160px; 
    font-size: 13px; 
    padding: 4px 4px 4px 4px; 
    border-radius: 2px; 
    -moz-border-radius: 2px; 
    -webkit-border-radius: 2px; 
    box-shadow: 0px 0px 2px #d9d9d9; 
    -moz-box-shadow: 0px 0px 2px #d9d9d9; 
    -webkit-box-shadow: 0px 0px 2px #d9d9d9; 
} 
 
.textboxtgl:focus { 
    outline: none; 
    border: 1px solid #7bc1f7; 
    box-shadow: 0px 0px 2px #7bc1f7; 
    -moz-box-shadow: 0px 0px 2px #7bc1f7; 
    -webkit-box-shadow: 0px 0px 2px #7bc1f7; 
} 
 </style>
 
<style>
.inp-icon{
background: url(cal4.png)no-repeat 100%;
background-size: 16px;
 }
</style>

<style>
.pagination {
  display: inline-block;
}

.pagination a {
  color: black;
  float: left;
  padding: 4px 12px;
  text-decoration: none;
  border: 1px solid #ddd;
  border-radius: 3px;
  font-size: 14px;
}

.pagination a.active {
  background-color: #C5C5E2;
  color: white;
  border: 1px solid #C5C5E2;
}

.pagination a:hover:not(.active) {background-color: #ddd;}

.pagination a:first-child {
  border-top-left-radius: 5px;
  border-bottom-left-radius: 5px;
}

.pagination a:last-child {
  border-top-right-radius: 5px;
  border-bottom-right-radius: 5px;
}
</style>

<style>

<style>

#logoleft {
width: 300px;
height: 120px;
position: absolute;
top: 40px;
left: 40px;
}

</style>

<script>
function addkct()
{
	document.getElementById('kctbarutab').style.display = 'block';
	//document.getElementById('barukct').style.display = 'none';
}
</script>

<script>
function clearparam()
{
	document.getElementsByName("namanpp")[0].value = "";
}
</script>

<script>
function viewpage(nohal)
{
var myval = document.getElementsByName("namanpp")[0].value;
if (myval==null || myval=="")
{
	var halink = "main.aspx?page=" + nohal + "&cari=none";
}
else
{
	var halink = "main.aspx?page=" + nohal + "&cari=" + myval;
}
//window.open(halink);
window.location.href = halink;
	//document.getElementsByName("pageno")[0].value = nohal;
	//document.forms["pagefrm"].submit();
}
</script>

<script>
function viewfilter()
{
var myval = document.getElementsByName("namanpp")[0].value;
var halink = "main.aspx?page=1&cari=" + myval;
//window.open(halink);
window.location.href = halink;
}
</script>

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
	Dim user0 as string
	Dim pwd0 as string
	Dim namauserx as String

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
	Dim userid as Integer = 1
	Dim jmlc as Integer = 0

	Dim SQL0 As String
	Dim SQL1 As String
	Dim SQL2 As String
	Dim nama20 as String = ""
	Dim posisi20 as String = ""
	
	Dim idq As String = ""
	Dim namauserq As String = ""
	Dim nppq as String = ""
	Dim statusq As String = ""
	Dim levelaksesq As String = ""
	Dim klpq As String = ""
	Dim unitq As String = ""
	Dim posisiq As String = ""
	Dim suratq As String = ""
	Dim sekreq As String = ""
	Dim divq As String = ""
	Dim adaq As String = ""
	Dim asetq As String = ""
	Dim statusaktifq As String = ""
	Dim posisix as String = ""
	Dim userp(1000) as String
	
	Dim statusmain as Integer = 0
	
	Dim aksesx as String = ""
	
	Dim setmenu as String = "utama"

	'lbl1.Text = Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)

	
	user0 = ""
	pwd0 = ""
	
	
    
	If NOT Session("username") is nothing Then
		statusmain = 1
		user0 = Session("username")
		idq = Session("userid")
		namauserq = Session("namauser")
		nppq = Session("idemployee")
		levelaksesq = Session("levelakses")
		klpq = Session("klp")
		unitq = Session("unit")
		posisiq = Session("posisi")
		posisix = Session("posisistr")
		suratq = Session("surat")
		divq = Session("divisi")
		suratq = Session("menusurat")
		sekreq = Session("menusekre")
		adaq = Session("menuada")
		asetq = Session("menuaset")
		statusaktifq = Session("statusaktif")
		levelaksesq = Session("levelakses")
	End If

	If NOT Request.Form("username") is nothing Then
		user0 = Request.Form("username")
		statusmain = 0
	End If

	If NOT Request.Form("passwd") is nothing Then
		pwd0 = Request.Form("passwd")
		pwd0 = pwd0.Replace("'","")
	End If

	
	aaa=1
	
    Try
        If (aaa=1)
			'myConn = New SqlConnection("Data Source=DESKTOP-S7N8DF9;Initial Catalog=KCT;User ID=sa;Password=sa;")
%>

<!--#include file="koneksi.aspx"-->

<%			
			
			myCmd = myConn.CreateCommand

'** Baca semua data sesuai Query			
			If statusmain = 0 Then
				SQL0 = "SELECT * FROM [user] WHERE (idemployee='" & user0 & "' AND paswd=HASHBYTES('SHA2_256','" & pwd0 & "'))"
				'SQL0 = "SELECT * FROM [user] WHERE (idemployee='" & user0 & "')"

				myCmd.CommandText = SQL0
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				
				a = 0
				idq = "P"
				SQL1 = "select"
				If myReader.HasRows Then
					a = 1
					While myReader.Read()
						ii = 1
						idq = myReader("ID").ToString()
						namauserq = myReader("nama").ToString()
						nppq = myReader("idemployee").ToString()
						statusq = myReader("status").ToString()
						levelaksesq = myReader("levelakses").ToString()
						klpq = myReader("klp").ToString()
						unitq = myReader("unit").ToString()
						posisiq = myReader("posisi").ToString()
						divq = myReader("div").ToString()
						suratq = myReader("menusurat").ToString()
						sekreq = myReader("menusekretaris").ToString()
						adaq = myReader("menupengadaan").ToString()
						asetq = myReader("menuaset").ToString()
						statusaktifq = myReader("status").ToString()
					End While
				End If
				myConn.Close()

			
				
				If a=0 OR ii=0 Then
%>
<script>
alert('Username/password anda salah. Silahkan signin lagi');
window.top.location.href = "logout.aspx"; 
</script>
<%
				End If
			End If

			If posisix.length()<5 Then
				SQL1 = "SELECT * FROM [posisi] WHERE (po_id=" & posisiq & ")"
				
				myCmd.CommandText = SQL1
				myConn.Open()
				myReader = myCmd.ExecuteReader()
				
				If myReader.HasRows Then
					While myReader.Read()
						posisix = myReader("posisi").ToString()
						posisi20 = posisix
					End While
				End If
				myConn.Close()
			End If
'========================================================================

			SQL2 = "SELECT * FROM [user]"

			myCmd.CommandText = SQL2
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			c = 0
			If myReader.HasRows Then
				While myReader.Read()
					c = c + 1
					userp(c) = myReader("idemployee").ToString()
				End While
			End If
			myConn.Close()
			jmlc = c
		
			
'========================================================================	

			If namauserq.length()<2 Then
%>

<script>
alert('Anda tidak beraktivitas terlalu lama (idle). Silahkan signin lagi');
window.top.location.href = "logout.aspx"; 
</script>

<%			
			End If


        End If

    Catch ex As Exception
        'lbl1.Text="Error while connecting Server."
		'Response.Write (ex.Message)

    Finally
		'myReader.Close()
        'myConn.Close() 'Whether there is error or not. Close the connection.

    End Try

	Dim strlength As Integer
	Dim strArr() As String
	Dim strnama1 as String

	strnama1 = namauserq
	strlength = namauserq.Length()

	If strlength > 20 Then
		strArr = namauserq.Split(" ")
		strnama1 = strArr(0) & " " & strArr(1)
		If strnama1.length() < 20 Then
			strnama1 = strArr(0) & " " & strArr(1) & " " & Ucase(strArr(2).substring(0,1))
		ElseIf strnama1.length() > 20 Then
			strnama1 = strArr(0) & " " & Ucase(strArr(1).substring(0,1))
		End If
	End If
	
	namauserq = strnama1
	Session("namauser") = namauserq
	
	If posisix.length() > 20 Then
		posisi20 = posisix.substring(0,20)
	End If
	
%>
<div class="example-element"></div>

<!--#include file="menu.aspx"-->

<!--#include file="topbar.aspx"-->

<div class="container">
			<div class="logo"><img id = "logoleft" src= "./images/portalbumlogo_red.png" style="width:180px;height:90px;"/></div>
			<div class="welcome">Selamat Datang,</div>
			<div class="photo"><img src="./images/user5.png" alt="" style="width:40px;height:40px;"></div>
			<div class="userin"><i><b><% Response.Write(namauserq) %></b></i></div>
			<div class="posisi"  STYLE="font-size: 0.8em"><i><% Response.Write(posisi20) %></i></div>
</div>

<table style="text-align: center; margin-left: 280px; margin-right: auto; margin-top: 0px;" width="75%" border="0" cellspacing="0" cellpadding="0">
	<tr height="20px">
		<td colspan="2" style="padding-right: 0px; text-align: right; font-size: 1.8em">
			<img src="./images/bnilogo2.png" alt="" style="width:110px;height:40px;">
		</td>
	</tr>
	<tr height="30px">
		<td colspan="2" style="padding-left: 0px; text-align: left; font-size: 1.8em">
			<FONT face="Arial" color="#000"><b></b></FONT>
		</td>
	</tr>
</table>
<%
'response.write("-----------------------------------++" & levelaksesq & "+" & SQL1 & "+" & nppq & "+" & ii)
'response.write("---------------------------------------++" & statusmain & "^^" & SQL2 & jmlc)
				For k=1 to jmlc
					SQL2 = "UPDATE [user] SET paswd=HASHBYTES('SHA2_256','" & userp(k) & "') WHERE idemployee='" & userp(k) & "'"
'response.write(SQL2)
					myCmd.CommandText = SQL2
					myConn.Open()
					myReader = myCmd.ExecuteReader()
					myConn.Close()
				Next k
%>
<table style="text-align: left; margin-left: 280px; margin-right: auto; margin-top: auto;" width="75%" border="0" cellspacing="0" cellpadding="0">
	<tr height="20px">
		<td style="padding-left: 0px; text-align: left; font-size: 1.8em" width="610px">
			<FONT face="Arial" color="#000">RESET &nbsp;<b><i>PORTAL BUM</i></b></FONT>
		</td>
		<td style="padding-left: 0px; text-align: left; font-size: 1.8em" width="340px">
			<FONT face="Arial" color="#000"></FONT>
		</td>
	</tr>
	<tr height="20px">
		<td colspan="2" style="padding-left: 0px; text-align: left; font-size: 1.8em">
			<FONT face="Arial" color="#000"><b></b></FONT>
		</td>
	</tr>
</table>
<table style="text-align: center; margin-left: 280px; margin-right: auto; margin-top: auto;" width="75%" border="0" cellspacing="0" cellpadding="0">
	<tr height="100px">
		<td colspan="2" style="padding-left: 0px; text-align: left; font-size: 1.0em">
			<img src="./images/letter2.jpg" alt="" style="width:100%;height:400px;">
		</td>
	</tr>
	<tr height="10px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top" width="140px">						
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top" width="40px">										
		</td>
	</tr>
	<tr height="20px">
		<td colspan="2" style="padding-left: 20px; text-align: left;">
		</td>
	</tr>
</table>

<%
		Session("username") = user0
		Session("userid") = idq
		Session("namauser") = namauserq
		Session("idemployee") = nppq
		Session("levelakses") = levelaksesq
		Session("klp") = klpq
		Session("unit") = unitq
		Session("posisi") = posisiq
		Session("posisistr") = posisix
		Session("surat") = suratq
		Session("divisi") = divq
		Session("menusurat") = suratq
		Session("menusekre") = sekreq 
		Session("menuada") = adaq
		Session("menuaset") = asetq
		Session("statusaktif") = statusaktifq
%>

<script>
$(document).ready(function(){
	$(".hamburger .hamburger__inner").click(function () {
	$(".wrapper").toggleClass("active")
})
 
$(".top_navbar .fas").click(function () {
	$(".profile_dd").toggleClass("active");
})
});
</script>


<%
'response.write("-------------------------------Default Timeout is: " & Session.Timeout)
%>




</body>
</html>
