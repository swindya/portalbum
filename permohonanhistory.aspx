<!DOCTYPE html>

<%@ Page Language="VB" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.Sql" %>
<%@ Import Namespace="System.Data.OleDb" %> 
<%@ Import Namespace="System.Configuration" %>
<%@ Import Namespace="System.Web.Configuration" %>

<%@ Import Namespace="ClosedXML.Excel" %>


<html lang="en">


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

<!--link rel="stylesheet" href="css1/styles.css"-->

<link href="./datepicker4/dcalendar.picker.css" rel="stylesheet" type="text/css">

<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	
<link href="./fontawesome/css/all.css" rel="stylesheet">


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
padding: 6px 14px 5px;
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
padding: 10px 14px 10px;
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
    height: 30px; 
    width: 160px; 
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
    height: 30px; 
    width: 370px; 
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
    height: 30px; 
    width: 650px; 
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
    width: 110px; 
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
select#bank {
    max-width: 350px;
    min-width: 220px;
    width: 220px !important;
}
</style>
<style>
select#jenis {
    max-width: 250px;
    min-width: 120px;
    width: 150px !important;
}
</style>

<style>
.select_style 
{
	background: #FFF;
	overflow: hidden;
	display: inline-block;
	color: #525252;
	font-weight: 300;
	-webkit-border-radius: 5px 4px 4px 5px/5px 5px 4px 4px;
	-moz-border-radius: 5px 4px 4px 5px/5px 5px 4px 4px;
	border-radius: 5px 4px 4px 5px/5px 5px 4px 4px;
	-webkit-box-shadow: 0 0 5px rgba(123, 123, 123, 0.2);
	-moz-box-shadow: 0 0 5px rgba(123,123,123,.2);
	box-shadow: 0 0 5px rgba(123, 123, 123, 0.2);
	border: solid 1px #DADADA;
	font-family: "helvetica neue",arial;
	position: relative;
	cursor: pointer;
	padding-right:2px;

}
.select_style span
{
	position: absolute;
	right: 5px;
	width: 8px;
	height: 8px;
	background: url ('./images/arrowdown1.png') no-repeat;
	top: 50%;
	margin-top: -4px;
}
.select_style select
{

	appearance:none;
	width:110%;
	background:none;
	background:transparent;
	border:none;
	outline:none;
	cursor:pointer;
	padding:4px 6px;
background: url('/images/arrowdown1.png') no-repeat 100% 4px #fff; /* add your own arrow image */
background: url('/images/arrowdown1.png') no-repeat 100% 4px #fff; /* fallback bg image*/
background: url('/images/arrowdown1.png') no-repeat 100% 0px, -webkit-linear-gradient(top, #fff, #9c9c9c);
background: url('/images/arrowdown1.png') no-repeat 100% 0px, -moz-linear-gradient(top, #fff, #9c9c9c);
background: url('/images/arrowdown1.png') no-repeat 100% 0px, -ms-linear-gradient(top, #fff, #9c9c9c);
background: url('/images/arrowdown1.png') no-repeat 100% 0px, -o-linear-gradient(top, #fff, #9c9c9c);
background: url('/images/arrowdown1.png') no-repeat 100% 0px, linear-gradient(top, #fff, #9c9c9c);
-webkit-appearance: menulist-button;
}
</style>

<style>
/* -------------------- Page Styles (not required) */
div { margin: 0px; }

/* -------------------- Select Box Styles: bavotasan.com Method (with special adaptations by ericrasch.com) */
/* -------------------- Source: http://bavotasan.com/2011/style-select-box-using-only-css/ */
.styled-select {
   background: url(15xvbd5.png) no-repeat 96% 0;
   height: 26px;
   overflow: hidden;
   width: 240px;
}

.styled-select select {
   background: transparent;
   border: none;
   font-size: 14px;
   height: 26px;
   padding: 5px; /* If you add too much padding here, the options won't show in IE */
   width: 268px;
}

.styled-select.slate {
   background: url(http://i62.tinypic.com/2e3ybe1.jpg) no-repeat right center;
   height: 30px;
   width: 240px;
}

.styled-select.slate select {
   border: 1px solid #ccc;
   font-size: 16px;
   height: 30px;
   width: 268px;
}

/* -------------------- Rounded Corners */
.rounded {
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
}

.semi-square {
   -webkit-border-radius: 5px;
   -moz-border-radius: 5px;
   border-radius: 5px;
}

/* -------------------- Colors: Background */
.slate   { background-color: #ddd; }
.green   { background-color: #779126; }
.blue    { background-color: #3b8ec2; }
.yellow  { background-color: #eec111; }
.black   { background-color: #000; }

/* -------------------- Colors: Text */
.slate select   { color: #000; }
.green select   { color: #fff; }
.blue select    { color: #fff; }
.yellow select  { color: #000; }
.black select   { color: #fff; }


/* -------------------- Select Box Styles: danielneumann.com Method */
/* -------------------- Source: http://danielneumann.com/blog/how-to-style-dropdown-with-css-only/ */
#mainselection select {
   border: 0;
   color: #EEE;
   background: transparent;
   font-size: 20px;
   font-weight: bold;
   padding: 2px 10px;
   width: 378px;
   *width: 350px;
   *background: #58B14C;
   -webkit-appearance: none;
}

#mainselection {
   overflow:hidden;
   width:350px;
   -moz-border-radius: 9px 9px 9px 9px;
   -webkit-border-radius: 9px 9px 9px 9px;
   border-radius: 9px 9px 9px 9px;
   box-shadow: 1px 1px 11px #330033;
   background: #58B14C url("15xvbd5.png") no-repeat scroll 319px center;
}


/* -------------------- Select Box Styles: stackoverflow.com Method */
/* -------------------- Source: http://stackoverflow.com/a/5809186 */
select#soflow, select#soflow-color {
   -webkit-appearance: button;
   -webkit-border-radius: 2px;
   -webkit-box-shadow: 0px 1px 3px rgba(0, 0, 0, 0.1);
   -webkit-padding-end: 20px;
   -webkit-padding-start: 2px;
   -webkit-user-select: none;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#FAFAFA, #F4F4F4 40%, #E5E5E5);
   background-position: 99% center;
   background-repeat: no-repeat;
   border: 1px solid #AAA;
   color: #555;
   font-size: 13px;
   margin: 0px;
   overflow: hidden;
   padding: 5px 10px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 150px;
}

select#soflow-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 15px;
}
select#soflow1, select#soflow-color {
   -webkit-appearance: button;
   -webkit-border-radius: 2px;
   -webkit-box-shadow: 0px 1px 3px rgba(0, 0, 0, 0.1);
   -webkit-padding-end: 20px;
   -webkit-padding-start: 2px;
   -webkit-user-select: none;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#FAFAFA, #F4F4F4 40%, #E5E5E5);
   background-position: 99% center;
   background-repeat: no-repeat;
   border: 1px solid #AAA;
   color: #555;
   font-size: 13px;
   margin: 0px;
   overflow: hidden;
   padding: 5px 10px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 300px;
}

select#soflow1-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 15px;
}
select#soflow2, select#soflow-color {
   -webkit-appearance: button;
   -webkit-border-radius: 2px;
   -webkit-box-shadow: 0px 1px 3px rgba(0, 0, 0, 0.1);
   -webkit-padding-end: 20px;
   -webkit-padding-start: 2px;
   -webkit-user-select: none;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#FAFAFA, #F4F4F4 40%, #E5E5E5);
   background-position: 99% center;
   background-repeat: no-repeat;
   border: 1px solid #AAA;
   color: #555;
   font-size: 13px;
   margin: 0px;
   overflow: hidden;
   padding: 5px 10px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 200px;
}

select#soflow2-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 15px;
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
html {
  //background: #f9f9f9;
  -webkit-box-sizing: border-box -moz-box-sizing: border-box;
  box-sizing: border-box;
}
a{
  text-decoration : none;
  color:#000;
}
.center-align {
    width: 100%;
    background-size: cover;
    margin: 0;
	//background:#232f3d;
    height: 0px;
	background-repeat: no-repeat;
}

.container {
 position: absolute;
    width: 100%;
    height: 100%;
}

footer {
  position: fixed;
  text-align: center;
  width: 100%;
  bottom: 10px;
}
footer a {
  padding-left:10px;
  padding-right:10px;
}
.search-bar{
  width: 100%;
    margin-top: 10px;
    font-size: 13px;
    padding: 10px;
    border: 1px solid #ccc;
}

.search-container {
    width: 474px;
    display: block;
    margin: 5px 30px;
    position: absolute;
    left: 1050px;
    top: 100px;
}
input#search-bar {
    margin: 0 auto;
    width: 100%;
    height: 32px;
    padding: 0 20px;
    font-size: 1.0rem;
    border: 1px solid #D0CFCE;
    outline: none;
    padding-right: 50px;
    margin-top: 23px;
    margin-left: -42px;
}
input#search-bar:focus {
  border: 1px solid #008ABF;
  transition: 0.35s ease;
  color: #000;
}
input#search-bar:focus::-webkit-input-placeholder {
  transition: opacity 0.45s ease;
  opacity: 0;
}
input#search-bar:focus::-moz-placeholder {
  transition: opacity 0.45s ease;
  opacity: 0;
}
input#search-bar:focus:-ms-placeholder {
  transition: opacity 0.45s ease;
  opacity: 0;
}

.search-icon {
    position: relative;
    float: right;
    width: 65px;
    height: 65px;
    top: -50px;
    right: -40px;
	left: -30px;
}


.spell-videos{
  display: none;
}
.final-image{
  width:100%;
  display: none;
}
.intro-text{
  text-align: center;
    width: 100%;
    font-weight: 200;
    font-size: 23px;
}
.cet{
  text-align:center;
  width:100%;
}
</style>

<style>
row.edit td { background-color: #CCEFFC; }
</style>

<style>
.chosen-container.chosen-container-single {
    width: 350px !important; /* or any value that fits your needs */
}
</style>
<style>
select, option {
width: 350px;
max-width: 350px;
}
</style>

<style>
.inp-icon{
background: url(cal4.png)no-repeat 100%;
background-size: 16px;
 }
</style>

<script>
function formatnumber(nField) {
//alert(this);
  //el.value = el.value.replace(/[^\d]/g,'').replace(/(\d\d?)$/,'.$1');
		if (/^0/.test(nField.value))
			{
			 nField.value = nField.value.substring(0,1);
			}
		if (Number(nField.value.replace(/,/g,"")))
			{
			 var tmp = nField.value.replace(/,/g,"");
	 		 tmp = tmp.toString().split('').reverse().join('').replace(/(\d{3})/g,'$1,').split('').reverse().join('').replace(/^,/,'');
			 if (/\./g.test(tmp))
				{
			 	 tmp = tmp.split(".");
				 tmp[1] = tmp[1].replace(/\,/g,"").replace(/ /,"");
				 nField.value = tmp[0]+"."+tmp[1]
				}
			 else 	{
				 nField.value = tmp.replace(/ /,"");
				} 
			}
		else	{
			 nField.value = nField.value.replace(/[^\d\,\.]/g,"").replace(/ /,"");
			}
}
</script>

<script>
function formatnilai(nilai) {

  //el.value = el.value.replace(/[^\d]/g,'').replace(/(\d\d?)$/,'.$1');
	var aa = parseInt(nilai);
	var aastr = aa.toLocaleString();
	return aastr;
}
</script>

<script>
function viewdata()
{
	document.forms["viewdatafrm"].submit();
}

</script>



<script>			

function vieweditkgd(id,tahun,tgl,jenisbelanja,pemohon,judul,step,filemohon,filespek,filetor,filerab,foldermohon,folderspek,foldertor,folderrab,nilai,kegiatan,kodekegiatan,statusedit)
{	
	document.getElementById('kpedittab').style.display = 'block';
	document.getElementById('kpbarutab').style.display = 'none';
	document.getElementById('barukp').style.display = 'none';
	document.getElementById('paramtab').style.display = 'none';
	
	if (statusedit==1)
	{
		document.getElementById('simpanbtn').style.display = 'none';
	}
	else
	{
		document.getElementById('simpanbtn').style.display = 'block';
	}

	var nila = formatnilai(nilai);
	document.getElementsByName("editgdid")[0].value = id;
	document.getElementsByName("edittgl")[0].value = tgl;
	document.getElementsByName("editjudul")[0].value = judul;
	document.getElementsByName("editnilai")[0].value = nilai.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
	document.getElementsByName("edittahun")[0].value = tahun;
	document.getElementsByName("editkegiatan")[0].value = kegiatan;
	document.getElementsByName("editkodekegiatan")[0].value = kodekegiatan;
	

	document.getElementById("editpemohon").value = pemohon;
	$("#editpemohon").val(pemohon).trigger("change");

	document.getElementById("editjenisbelanja").value = jenisbelanja;
	$("#editjenisbelanja").val(jenisbelanja).trigger("change");
	
	//document.getElementById("editkeputusan").value = keputusan;
	//$("#editkeputusan").val(keputusan).trigger("change");

//--------------------------------------------------------------------------------
	if (filemohon.length > 2)
	{
		var namafilearr = filemohon.split(".");
		var noarr = namafilearr.length;
		var jenisfile = namafilearr[noarr-1];

		if (jenisfile==='pdf' || jenisfile.length > 4)
		{       
			document.getElementById("edituploadPreview1").src = './images/pdficon1.png'; 	
		}
		else
		{           
			document.getElementById("edituploadPreview1").src = "."+foldermohon+filemohon; 		
		}
		document.getElementById("editlinkimage1").href = "."+foldermohon+filemohon;
	}
	else
	{
		document.getElementById("edituploadPreview1").src = './images/nofile.png';
	}
//--------------------------------------------------------------------------------	
}

</script>	

<script>
function editcuti()
{
	//document.getElementById('cutiedittab').style.display = 'block';
}
</script>

<script>
function hideeditkp()
{
	window.opener.location.reload();
	window.close();
}
</script>

<script>
function getpegdata(nip)
{
	//alert(nip);
}
</script>

<script>
function viewhps(userid, adaid)
{
	
	var textfield = document.createElement("input");
	textfield.type = "hidden";
	textfield.name = "userid";
	textfield.id = "userid";
	textfield.value = userid;
	var textfield1 = document.createElement("input");
	textfield1.type = "hidden";
	textfield1.name = "adaid";
	textfield1.id = "adaid";
	textfield1.value = adaid;
	var textfield2 = document.createElement("input");
	textfield2.type = "hidden";
	textfield2.name = "statusedit";
	textfield2.id = "statusedit";
	textfield2.value = 2;
	document.getElementById('viewhpsfrm').appendChild(textfield);
	document.getElementById('viewhpsfrm').appendChild(textfield1);
	document.getElementById('viewhpsfrm').appendChild(textfield2);
	document.forms["viewhpsfrm"].submit();
}
</script>

<script>
function viewundangan(userid, adaid)
{
	
	var textfield = document.createElement("input");
	textfield.type = "hidden";
	textfield.name = "userid";
	textfield.id = "userid";
	textfield.value = userid;
	var textfield1 = document.createElement("input");
	textfield1.type = "hidden";
	textfield1.name = "adaid";
	textfield1.id = "adaid";
	textfield1.value = adaid;
	var textfield2 = document.createElement("input");
	textfield2.type = "hidden";
	textfield2.name = "statusedit";
	textfield2.id = "statusedit";
	textfield2.value = 2;
	document.getElementById('viewundanganfrm').appendChild(textfield);
	document.getElementById('viewundanganfrm').appendChild(textfield1);
	document.getElementById('viewundanganfrm').appendChild(textfield2);
	document.forms["viewundanganfrm"].submit();
}
</script>

<script>
function viewpenawaran(userid, adaid)
{
	
	var textfield = document.createElement("input");
	textfield.type = "hidden";
	textfield.name = "userid";
	textfield.id = "userid";
	textfield.value = userid;
	var textfield1 = document.createElement("input");
	textfield1.type = "hidden";
	textfield1.name = "adaid";
	textfield1.id = "adaid";
	textfield1.value = adaid;
	var textfield2 = document.createElement("input");
	textfield2.type = "hidden";
	textfield2.name = "statusedit";
	textfield2.id = "statusedit";
	textfield2.value = 2;
	document.getElementById('viewpenawaranfrm').appendChild(textfield);
	document.getElementById('viewpenawaranfrm').appendChild(textfield1);
	document.getElementById('viewpenawaranfrm').appendChild(textfield2);
	document.forms["viewpenawaranfrm"].submit();
}
</script>

<script>
function viewnego(userid, adaid)
{
	
	var textfield = document.createElement("input");
	textfield.type = "hidden";
	textfield.name = "userid";
	textfield.id = "userid";
	textfield.value = userid;
	var textfield1 = document.createElement("input");
	textfield1.type = "hidden";
	textfield1.name = "adaid";
	textfield1.id = "adaid";
	textfield1.value = adaid;
	var textfield2 = document.createElement("input");
	textfield2.type = "hidden";
	textfield2.name = "statusedit";
	textfield2.id = "statusedit";
	textfield2.value = 2;
	document.getElementById('viewnegosiasifrm').appendChild(textfield);
	document.getElementById('viewnegosiasifrm').appendChild(textfield1);
	document.getElementById('viewnegosiasifrm').appendChild(textfield2);
	document.forms["viewnegosiasifrm"].submit();
}
</script>

<script>
function viewhasil(userid, adaid)
{
	
	var textfield = document.createElement("input");
	textfield.type = "hidden";
	textfield.name = "userid";
	textfield.id = "userid";
	textfield.value = userid;
	var textfield1 = document.createElement("input");
	textfield1.type = "hidden";
	textfield1.name = "adaid";
	textfield1.id = "adaid";
	textfield1.value = adaid;
	var textfield2 = document.createElement("input");
	textfield2.type = "hidden";
	textfield2.name = "statusedit";
	textfield2.id = "statusedit";
	textfield2.value = 2;
	document.getElementById('viewhasilfrm').appendChild(textfield);
	document.getElementById('viewhasilfrm').appendChild(textfield1);
	document.getElementById('viewhasilfrm').appendChild(textfield2);
	document.forms["viewhasilfrm"].submit();
}
</script>

<script>
function viewpenunjukan(userid, adaid)
{
	var textfield = document.createElement("input");
	textfield.type = "hidden";
	textfield.name = "userid";
	textfield.id = "userid";
	textfield.value = userid;
	var textfield1 = document.createElement("input");
	textfield1.type = "hidden";
	textfield1.name = "adaid";
	textfield1.id = "adaid";
	textfield1.value = adaid;
	var textfield2 = document.createElement("input");
	textfield2.type = "hidden";
	textfield2.name = "statusedit";
	textfield2.id = "statusedit";
	textfield2.value = 2;
	document.getElementById('viewtunjukfrm').appendChild(textfield);
	document.getElementById('viewtunjukfrm').appendChild(textfield1);
	document.getElementById('viewtunjukfrm').appendChild(textfield2);
	document.forms["viewtunjukfrm"].submit();
}
</script>

<script>
function viewpesanan(userid, adaid)
{
	
	var textfield = document.createElement("input");
	textfield.type = "hidden";
	textfield.name = "userid";
	textfield.id = "userid";
	textfield.value = userid;
	var textfield1 = document.createElement("input");
	textfield1.type = "hidden";
	textfield1.name = "adaid";
	textfield1.id = "adaid";
	textfield1.value = adaid;
	var textfield2 = document.createElement("input");
	textfield2.type = "hidden";
	textfield2.name = "statusedit";
	textfield2.id = "statusedit";
	textfield2.value = 2;
	document.getElementById('viewpesanfrm').appendChild(textfield);
	document.getElementById('viewpesanfrm').appendChild(textfield1);
	document.getElementById('viewpesanfrm').appendChild(textfield2);
	document.forms["viewpesanfrm"].submit();
}
</script>

<script>
function viewspk(userid, adaid)
{
	
	var textfield = document.createElement("input");
	textfield.type = "hidden";
	textfield.name = "userid";
	textfield.id = "userid";
	textfield.value = userid;
	var textfield1 = document.createElement("input");
	textfield1.type = "hidden";
	textfield1.name = "adaid";
	textfield1.id = "adaid";
	textfield1.value = adaid;
	var textfield2 = document.createElement("input");
	textfield2.type = "hidden";
	textfield2.name = "statusedit";
	textfield2.id = "statusedit";
	textfield2.value = 2;
	document.getElementById('viewspkfrm').appendChild(textfield);
	document.getElementById('viewspkfrm').appendChild(textfield1);
	document.getElementById('viewspkfrm').appendChild(textfield2);
	document.forms["viewspkfrm"].submit();
}
</script>

<script>
function viewspp(userid, adaid)
{
	
	var textfield = document.createElement("input");
	textfield.type = "hidden";
	textfield.name = "userid";
	textfield.id = "userid";
	textfield.value = userid;
	var textfield1 = document.createElement("input");
	textfield1.type = "hidden";
	textfield1.name = "adaid";
	textfield1.id = "adaid";
	textfield1.value = adaid;
	var textfield2 = document.createElement("input");
	textfield2.type = "hidden";
	textfield2.name = "statusedit";
	textfield2.id = "statusedit";
	textfield2.value = 2;
	document.getElementById('viewsppfrm').appendChild(textfield);
	document.getElementById('viewsppfrm').appendChild(textfield1);
	document.getElementById('viewsppfrm').appendChild(textfield2);
	document.forms["viewsppfrm"].submit();
}
</script>

<script>
function viewpks(userid, adaid)
{
	
	var textfield = document.createElement("input");
	textfield.type = "hidden";
	textfield.name = "userid";
	textfield.id = "userid";
	textfield.value = userid;
	var textfield1 = document.createElement("input");
	textfield1.type = "hidden";
	textfield1.name = "adaid";
	textfield1.id = "adaid";
	textfield1.value = adaid;
	var textfield2 = document.createElement("input");
	textfield2.type = "hidden";
	textfield2.name = "statusedit";
	textfield2.id = "statusedit";
	textfield2.value = 2;
	document.getElementById('viewpksfrm').appendChild(textfield);
	document.getElementById('viewpksfrm').appendChild(textfield1);
	document.getElementById('viewpksfrm').appendChild(textfield2);
	document.forms["viewpksfrm"].submit();
}
</script>

<script>
function viewselesai(userid, adaid)
{
	
	var textfield = document.createElement("input");
	textfield.type = "hidden";
	textfield.name = "userid";
	textfield.id = "userid";
	textfield.value = userid;
	var textfield1 = document.createElement("input");
	textfield1.type = "hidden";
	textfield1.name = "adaid";
	textfield1.id = "adaid";
	textfield1.value = adaid;
	var textfield2 = document.createElement("input");
	textfield2.type = "hidden";
	textfield2.name = "statusedit";
	textfield2.id = "statusedit";
	textfield2.value = 2;
	document.getElementById('viewselesaifrm').appendChild(textfield);
	document.getElementById('viewselesaifrm').appendChild(textfield1);
	document.getElementById('viewselesaifrm').appendChild(textfield2);
	document.forms["viewselesaifrm"].submit();
}
</script>

<script>
function viewserahterima(userid, adaid)
{
	
	var textfield = document.createElement("input");
	textfield.type = "hidden";
	textfield.name = "userid";
	textfield.id = "userid";
	textfield.value = userid;
	var textfield1 = document.createElement("input");
	textfield1.type = "hidden";
	textfield1.name = "adaid";
	textfield1.id = "adaid";
	textfield1.value = adaid;
	var textfield2 = document.createElement("input");
	textfield2.type = "hidden";
	textfield2.name = "statusedit";
	textfield2.id = "statusedit";
	textfield2.value = 2;
	document.getElementById('viewserahterimafrm').appendChild(textfield);
	document.getElementById('viewserahterimafrm').appendChild(textfield1);
	document.getElementById('viewserahterimafrm').appendChild(textfield2);
	document.forms["viewserahterimafrm"].submit();
}
</script>

<script>
function viewperiksa(userid, adaid)
{
	
	var textfield = document.createElement("input");
	textfield.type = "hidden";
	textfield.name = "userid";
	textfield.id = "userid";
	textfield.value = userid;
	var textfield1 = document.createElement("input");
	textfield1.type = "hidden";
	textfield1.name = "adaid";
	textfield1.id = "adaid";
	textfield1.value = adaid;
	var textfield2 = document.createElement("input");
	textfield2.type = "hidden";
	textfield2.name = "statusedit";
	textfield2.id = "statusedit";
	textfield2.value = 2;
	document.getElementById('viewperiksafrm').appendChild(textfield);
	document.getElementById('viewperiksafrm').appendChild(textfield1);
	document.getElementById('viewperiksafrm').appendChild(textfield2);
	document.forms["viewperiksafrm"].submit();
}
</script>

<script>
function viewbayar(userid, adaid)
{
	
	var textfield = document.createElement("input");
	textfield.type = "hidden";
	textfield.name = "userid";
	textfield.id = "userid";
	textfield.value = userid;
	var textfield1 = document.createElement("input");
	textfield1.type = "hidden";
	textfield1.name = "adaid";
	textfield1.id = "adaid";
	textfield1.value = adaid;
	var textfield2 = document.createElement("input");
	textfield2.type = "hidden";
	textfield2.name = "statusedit";
	textfield2.id = "statusedit";
	textfield2.value = 2;
	document.getElementById('viewbayarfrm').appendChild(textfield);
	document.getElementById('viewbayarfrm').appendChild(textfield1);
	document.getElementById('viewbayarfrm').appendChild(textfield2);
	document.forms["viewbayarfrm"].submit();
}
</script>

<script>
function saveeditkp()
{
	document.forms["editkpfrm"].submit();
}
</script>

<script>
function viewijin(userid, adaid)
{
	document.getElementsByName('ijinuserid')[0].value = userid;
	document.getElementsByName('ijinadaid')[0].value = adaid;
	document.forms["viewijinfrm"].submit();
}
</script>

<script>
function viewusul(userid, adaid)
{
	document.getElementsByName('usuluserid')[0].value = userid;
	document.getElementsByName('usuladaid')[0].value = adaid;
	document.forms["viewusulfrm"].submit();
}
</script>

<script>
function viewpoc(userid, adaid)
{
	document.getElementsByName('pocuserid')[0].value = userid;
	document.getElementsByName('pocadaid')[0].value = adaid;
	document.forms["viewpocfrm"].submit();
}
</script>

<script>
function viewaanwij(userid, adaid)
{
	document.getElementsByName('aanuserid')[0].value = userid;
	document.getElementsByName('aanadaid')[0].value = adaid;
	document.forms["viewaanwijfrm"].submit();
}
</script>

<script>
function vieweval(userid, adaid)
{
	document.getElementsByName('evaluserid')[0].value = userid;
	document.getElementsByName('evaladaid')[0].value = adaid;
	document.forms["viewevalfrm"].submit();
}
</script>

<script>
function viewnego(userid, adaid)
{
	document.getElementsByName('aanuserid')[0].value = userid;
	document.getElementsByName('aanadaid')[0].value = adaid;
	document.forms["viewnegofrm"].submit();
}
</script>

<script>
function viewnkp(userid, adaid)
{
	document.getElementsByName('nkpuserid')[0].value = userid;
	document.getElementsByName('nkpadaid')[0].value = adaid;
	document.forms["viewnkpfrm"].submit();
}
</script>

<script>
function viewsanggah(userid, adaid)
{
	document.getElementsByName('sanggahuserid')[0].value = userid;
	document.getElementsByName('sanggahadaid')[0].value = adaid;
	document.forms["viewsanggahfrm"].submit();
}
</script>

<script>
function hapuskp(a)
{
	var r = confirm("Apakah anda yakin akan hapus data ?");
	if (r==true)
	{
		document.getElementsByName('delkpid')[0].value = a;
		document.forms["delkpfrm"].submit();
	}
}
</script>

<script>
function gotohasil()
{
	document.forms["gotohasilfrm"].submit();
}
</script>

<script type="text/javascript">     
function PreviewImage(no) 
{         
	var oFReader = new FileReader();
	var namafil = document.getElementById("uploadImage"+no).value;
	var namafilearr = namafil.split(".");
	var noarr = namafilearr.length;
	var jenisfile = namafilearr[noarr-1];

	oFReader.readAsDataURL(document.getElementById("uploadImage"+no).files[0]);
	
	if (jenisfile==='pdf')
	{
		oFReader.onload = function (oFREvent) {             
		document.getElementById("uploadPreview"+no).src = './images/pdficon1.png'; 	
		};
	}
	else
	{
		oFReader.onload = function (oFREvent) {             
		document.getElementById("uploadPreview"+no).src = oFREvent.target.result; 		
		};
	}
	//document.forms["memobarufrm"].submit();
	//memousulanbaru();

} 
</script>
<script type="text/javascript">     
function editPreviewImage(no) 
{         
	var oFReader = new FileReader();
	var namafil = document.getElementById("edituploadImage"+no).value;
	var namafilearr = namafil.split(".");
	var noarr = namafilearr.length;
	var jenisfile = namafilearr[noarr-1];

	oFReader.readAsDataURL(document.getElementById("edituploadImage"+no).files[0]);
	
	if (jenisfile==='pdf')
	{
		oFReader.onload = function (oFREvent) {             
		document.getElementById("edituploadPreview"+no).src = './images/pdficon1.png'; 	
		};
	}
	else
	{
		oFReader.onload = function (oFREvent) {             
		document.getElementById("edituploadPreview"+no).src = oFREvent.target.result; 		
		};
	}
	//document.forms["memobarufrm"].submit();
	//memousulanbaru();

} 
</script>

	
</head>

<body>
    <div class="page-wrapper default-theme sidebar-bg bg3 toggled">

<%
    'Inherits System.Windows.Forms.Form
    'Create ADO.NET objects.
    Dim myConn As SqlConnection
    Dim myCmd As SqlCommand
	Dim myCmdi As SqlCommand
    Dim myReader As SqlDataReader
    'results As String
	Dim a, aaa, b, bb, c, d, dd, jmla, jmldata as integer
	Dim user0 as string
	Dim pwd0 as string
	Dim namauserx as String
	Dim aksesx as String
	Dim namamohon as String = ""
	Dim idx as String
	Dim nosuratx as String
	Dim tglx as String
	Dim tahunx as String
	Dim jenisx as String
	Dim jenisangx as String = ""
	Dim namax as String
	Dim nilaix as String
	Dim nilaihpsx as String = ""
	Dim tembusanx as String
	Dim cuserx as String
	Dim nostepx as String
	Dim stepx as String
	Dim nostepijinx as String
	Dim statusrevisix as String
	Dim filenamex as String
	Dim filestr as String
	Dim statusf as Integer = 0
	Dim ketx as String = ""
	'Dim pageno as Integer = 1
	
	Dim rowperpage as Integer = 20
	Dim pageno as String
	Dim startpage as Long = 0
	Dim endpage as Long = 0
	Dim jmlpage as Integer = 1
	Dim pagerows as Integer = 20
	Dim rowsperpage as Integer = 20
	Dim hal as Integer = 1
	
	Dim useridx as Integer = 0
	Dim itop as Integer = 1
	Dim ii as Integer = 0
	Dim k as Integer
	Dim userid as Integer = 0
	
	Dim idz(1000) as String
	Dim idq As String = ""
	Dim namauserq As String = ""
	Dim statusq As String = ""
	Dim jabatanq As String = ""
	Dim klpq As String = ""
	Dim unitq As String = ""
	Dim posisiqq As String = ""
	Dim posisiq As Integer = 0
	Dim posisint as Integer = 0
	Dim suratq As String = ""
	Dim sekreq As String = ""
	Dim divq As String = ""
	Dim nppq as String = ""
	Dim levelaksesq as String = ""
	Dim adaq as String = ""
	Dim asetq as String = ""
	Dim statusaktifq as String = ""
	Dim posisix as String = ""
	Dim namastr as String = ""
	Dim tglstr as String = ""
	Dim tgl1 as String = ""
	Dim tgl2 as String = ""
	Dim tglqu as String = ""
	Dim tahap as String  = ""
	Dim tahapstr as String = ""
	Dim loguserstr as String = ""
	Dim adaid as Long = 0
	Dim tglspkx as String = ""

	Dim SQL0 As String
	Dim SQL1 As String
	Dim SQL2 As String
	Dim SQL3 as String
	Dim nama20 as String = ""
	Dim posisi20 as String = ""
	Dim unitstr as String = ""
	Dim klpstr as String = ""
	Dim divstr as String = ""
	Dim jenisfile1 as String = ""
	Dim jenisfile2 as String = ""
	Dim jenisfile3 as String = ""
	Dim jenisfile4 as String = ""
	Dim jenisfile5 as String = ""
	Dim jenisfile6 as String = ""
	Dim jenisarr() as String
	Dim jeniscount as Integer = 0
	
	Dim perihal as String = ""
	Dim perihalstr as String = ""
	Dim filestrx as String = ""
	Dim hpsx as Long = 0
	
	Dim startval as Integer = 0
	Dim endval as Integer = 0
	Dim datastrx, tglxx as String
	Dim tahuniki as Integer
	Dim tahunwingi as Integer
	Dim tahunsesuk as Integer
	Dim tgl10 as String = ""
	Dim tgl20 as String = ""
	Dim tglarr(10) as String
	Dim stepstr as String = ""
	Dim alasanx as String = ""
	Dim jmliji, jmlusu, jmlaan, jmlpoc, jmleva, jmlneg, jmlnkp, jmlsan, jmlpen, jmlspk, jmlpks as Integer
	Dim folderx, filesuratx, file1x, file2x, file3x, file4x, file5x, file6x, file7x, file8x as String
	Dim fileipx, filerabx, filetorx as String
	Dim tglawalx, tglakhirx as String
	Dim nostepv, nilaijml, nilaihps as Long
	Dim statusrejectx as Integer = 0
	Dim dtgl as Date
	
	Dim setmenu as String = "mohon"

	Dim regDate as Date = Date.Now()
	Dim strDate as String = regDate.ToString("dd-MM-yyyy")
	
	Dim nomory as String = ""
	Dim jmlprenegox as Long = 0
	Dim jmlpostnegox as Long = 0
	Dim jmlpostnegox2 as Long = 0
	Dim jmlpostnegox3 as Long = 0
	Dim jmlpostnegox4 as Long = 0
	Dim jmlpostnegox5 as Long = 0
	Dim jmlpostnegox6 as Long = 0
	Dim vendorx as String = ""
	Dim vendorx2 as String = ""
	Dim vendorx3 as String = ""
	Dim vendorx4 as String = ""
	Dim vendorx5 as String = ""
	Dim vendorx6 as String = ""
	Dim negostr as String = ""
	
	Dim vendor as String = ""
	Dim vendor2 as String = ""
	Dim vendor3 as String = ""
	Dim vendor4 as String = ""
	Dim vendor5 as String = ""
	Dim vendor6 as String = ""
	Dim vendorz(10) as String
	Dim vendorzz(10) as String
	Dim nomorz(10) as String
	Dim jmlnego as Long = 0
	Dim jmlnego2 as Long = 0
	Dim jmlnego3 as Long = 0
	Dim jmlnego4 as Long = 0
	Dim jmlnego5 as Long = 0
	Dim jmlnego6 as Long = 0
	Dim jmlnegoz(10) as Long
	Dim jmlnegozz(10) as Long
	Dim vendorstr as String = ""

	'lbl1.Text = Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
	
	tahuniki = Val(Year(Now))
	tahunwingi = tahuniki - 1
	tahunsesuk = tahuniki + 1

	jmliji = 0
	jmlusu = 0
	jmlaan = 0
	jmlpoc = 0
	jmleva = 0
	jmlneg = 0
	jmlnkp = 0
	jmlsan = 0
	jmlpen = 0
	jmlspk = 0
	jmlpks = 0


	aaa=1
	
	user0 = ""
	pwd0 = ""

%>
<!--#include file="koneksi.aspx"-->
<%

    Try
        If aaa=1 Then
			'myConn = New SqlConnection("Data Source=DESKTOP-S7N8DF9;Initial Catalog=KCT;User ID=sa;Password=sa;")

			If NOT Request.Form("hisuserid") is Nothing Then
				userid = Val(Request.Form("hisuserid"))
			End If
			If NOT Request.Form("hadaid") is Nothing Then
				adaid = Request.Form("hadaid")
			End If

			
			user0 = Session("username")
			idq = Session("userid")
			nppq = Session("idemployee")
			If not Session("namauser") is nothing Then
				namauserq = Session("namauser")
				nama20 = Left(namauserq, 16)
			End If
			jabatanq = Session("jabatan")
			klpq = Session("klp")
			unitq = Session("unit")
			posisiq = Val(Session("posisi"))
			posisix = Session("posisistr")
			suratq = Session("menusurat")
			sekreq = Session("menusekre")
			adaq = Session("menuada")
			asetq = Session("menuaset")
			divq = Session("divisi")
			levelaksesq = Session("levelakses")
			posisi20 = Left(posisix, 20)


			If userid=0 Then
				userid = idq
			End If

%>

<table style="text-align: left; margin-left: 50px; margin-right: auto; margin-top: auto;" width="1200px" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td style="padding-left: auto; text-align: left;">
			<table style="text-align: left; margin-left: auto; margin-right: auto; margin-top: auto;" width="1200px" border="0" cellspacing="0" cellpadding="0">
				<tr height="50px">
					<td colspan="2" style="padding-left: 0px; text-align: left;">
					</td>
				</tr>
				<tr height="30px">
					<!--td style="padding-left: 0px; text-align: left;" width="50px">
						<img src="images/homecircle1.png" alt="Girl in a jacket" style="width:33px;height:33px">
					</td-->
					<td style="padding-left: 0px; text-align: left; font-size: 1.5em" width="1150px">
						<FONT face="Arial" color="#222">KRONOLOGIS/HISTORICAL PENGADAAN</FONT>
					</td>
				</tr>
				<tr height="20px">
					<td style="padding-left: 10px; text-align: right;">
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>

<h2 class="center-align">
   <!--form class="search-container" onsubmit="return false;">
    <input type="text" id="search-bar" placeholder="Nama / NIP / Unit Kerja">
    <a href="#" id="spell-init"><img class="search-icon" src="images/search-icon.png"></a>
  </form-->
  
</h2>
  
<%

'#######################################################################################

			myCmd = myConn.CreateCommand

'==============================================================================
' 1. BACA DATA PENGADAAN - IJIN PRINSIP
'==============================================================================
			SQL3 = "SELECT * FROM pengadaan WHERE ID=" & adaid
			myCmd.CommandText = SQL3
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			If myReader.HasRows Then
				While myReader.Read()
					idx = myReader("ID").ToString()
					nosuratx = myReader("kode").ToString()
					tglx = myReader("tglmohon").ToString()
					tahunx = myReader("tahun").ToString()
					jenisx = myReader("jenispengadaan").ToString()
					jenisangx = myReader("jenisanggaran").ToString()
					namax = myReader("nama").ToString()
					nostepx = myReader("nostep").ToString()
					nostepv = Val(nostepx)
					statusrevisix = myReader("statusrevisi").ToString()
					statusrejectx = Val(myReader("statusreject").ToString())
					dtgl = Convert.ToDateTime(tglx)
					tglxx = dtgl.ToString("dd-MM-yyyy")
					'Dim oDate As DateTime = Convert.ToDateTime(tglx)
					jmliji = 1
				End While
			End If
			myConn.Close()

'==============================================================================
' 2. BACA DATA PROSES USULAN
'==============================================================================
			SQL3 = "SELECT * FROM pengadaan_prosesusulan WHERE pengadaanID=" & adaid
'response.write(SQL3 & "<br>")
			myCmd.CommandText = SQL3
			myConn.Open()
			myReader = myCmd.ExecuteReader()

			If myReader.HasRows Then
				While myReader.Read()
					jmlusu = 1
				End While
			End If
			myConn.Close()
'==============================================================================
' 3. BACA DATA AANWIJZING
'==============================================================================
			SQL3 = "SELECT * FROM pengadaan_aanwijzing WHERE pengadaanID=" & adaid
'response.write(SQL3 & "<br>")
			myCmd.CommandText = SQL3
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			If myReader.HasRows Then
				While myReader.Read()
					jmlaan = 1
				End While
			End If
			myConn.Close()
'==============================================================================
' 4. BACA DATA POC
'==============================================================================
			SQL3 = "SELECT * FROM pengadaan_poc WHERE pengadaanID=" & adaid

			myCmd.CommandText = SQL3
			myConn.Open()
			myReader = myCmd.ExecuteReader()

			If myReader.HasRows Then
				While myReader.Read()
					jmlpoc = 1
				End While
			End If
			myConn.Close()
'==============================================================================
' 5. BACA DATA EVALUASI ADMIN & TEKNIK
'==============================================================================
			SQL3 = "SELECT * FROM pengadaan_evaluasi WHERE pengadaanID=" & adaid
			myCmd.CommandText = SQL3
			myConn.Open()
			myReader = myCmd.ExecuteReader()

			If myReader.HasRows Then
				While myReader.Read()
					jmleva = 1
				End While
			End If
			myConn.Close()
'==============================================================================
' 6. BACA DATA NEGOSIASI
'==============================================================================
			SQL3 = "SELECT * FROM pengadaan_negosiasi WHERE pengadaanID=" & adaid

			myCmd.CommandText = SQL3
			myConn.Open()
			myReader = myCmd.ExecuteReader()

			If myReader.HasRows Then
				While myReader.Read()
					jmlneg = 1
				End While
			End If
			myConn.Close()
'==============================================================================
' 7. BACA DATA NKP
'==============================================================================
			SQL3 = "SELECT * FROM pengadaan_nkp WHERE pengadaanID=" & adaid
			myCmd.CommandText = SQL3
			myConn.Open()
			myReader = myCmd.ExecuteReader()

			If myReader.HasRows Then
				While myReader.Read()
					jmlnkp = 1
				End While
			End If
			myConn.Close()
'==============================================================================
' 8. BACA DATA SANGGAHAN
'==============================================================================
			SQL3 = "SELECT * FROM pengadaan_sanggah WHERE pengadaanID=" & adaid
			myCmd.CommandText = SQL3
			myConn.Open()
			myReader = myCmd.ExecuteReader()

			If myReader.HasRows Then
				While myReader.Read()
					jmlsan = 1
				End While
			End If
			myConn.Close()
'==============================================================================
' 9. BACA DATA PENUNJUKAN PEMENANG
'==============================================================================
			SQL3 = "SELECT * FROM pengadaan_penunjukan WHERE pengadaanID=" & adaid
			myCmd.CommandText = SQL3
			myConn.Open()
			myReader = myCmd.ExecuteReader()

			If myReader.HasRows Then
				While myReader.Read()
					jmlpen = 1
				End While
			End If
			myConn.Close()

'==============================================================================
' 10. BACA DATA SPK
'==============================================================================
			SQL3 = "SELECT * FROM pengadaan_spk WHERE pengadaanID=" & adaid
			myCmd.CommandText = SQL3
			myConn.Open()
			myReader = myCmd.ExecuteReader()

			If myReader.HasRows Then
				While myReader.Read()
					jmlspk = 1
				End While
			End If
			myConn.Close()

'==============================================================================
' 11. BACA DATA PKS
'==============================================================================
			SQL3 = "SELECT * FROM pengadaan_pks WHERE pengadaanID=" & adaid
			myCmd.CommandText = SQL3
			myConn.Open()
			myReader = myCmd.ExecuteReader()

			If myReader.HasRows Then
				While myReader.Read()
					jmlpks = 1
				End While
			End If
			myConn.Close()
'======================================================================================

d = 0
%>

<table id="kpedittab" style="text-align: left; margin-left: 50px; margin-right: auto; margin-top: auto;" border="0" cellspacing="0" cellpadding="0">
	<!--tr height="20px">
		<td colspan="3" style="padding-left: 0px; text-align: left; font-size: 1.5em">
			<FONT face="Arial" color="#000"><b></b></FONT>
			<img src="./images/edit2.jpg" alt="" style="width:80px;height:80px;">
		</td>
	</tr-->
	<tr height="50px">
		<td colspan="3" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		</td>
	</tr>
	<tr height="0px">
		<td colspan="3" style="padding-left: 20px; text-align: left; font-size: 1.0em">
			<input type="hidden" id="editmohonuserid" name="editmohonuserid" value="<?php echo $userid;?>">
		</td>
	</tr>
	<tr height="30px">
		<td style="padding-left: 0px; text-align: left; font-size: 1.1em; vertical-align: top" width="200px">
			<FONT face="Arial" color="#000">
				Tahun Anggaran
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.1em; vertical-align: top" width="30px">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; font-family: Arial; font-size: 1.1em; vertical-align: top" width="880px">
			<FONT face="Arial" color="#000">
<% 
			response.write(tahunx)
%>
			</FONT>		
		</td>
	</tr>
	<tr height="30px">
		<td style="padding-left: 0px; text-align: left; font-size: 1.1em; vertical-align: top" width="200px">
			<FONT face="Arial" color="#000">
				Nama Pengadaan
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.1em; vertical-align: top" width="30px">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; font-family: Arial; font-size: 1.1em; vertical-align: top" width="880px">
			<FONT face="Arial" color="#000">
<% 
			response.write(namax)
%>
			</FONT>		
		</td>
	</tr>
	<tr height="30px">
		<td style="padding-left: 0px; text-align: left; font-size: 1.1em; vertical-align: top" width="200px">
			<FONT face="Arial" color="#000">
				Jenis Pengadaan
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.1em; vertical-align: top" width="30px">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; font-family: Arial; font-size: 1.1em; vertical-align: top" width="880px">
			<FONT face="Arial" color="#000">
<%
			response.write(jenisx.ToUpper())
%>
			</FONT>		
		</td>
	</tr>
	<tr height="30px">
		<td style="padding-left: 0px; text-align: left; font-size: 1.1em; vertical-align: top" width="200px">
			<FONT face="Arial" color="#000">
				Jenis Anggaran
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.1em; vertical-align: top" width="30px">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; font-family: Arial; font-size: 1.1em; vertical-align: top" width="880px">
			<FONT face="Arial" color="#000">
<%
			response.write(jenisangx.ToUpper())
%>
			</FONT>		
		</td>
	</tr>
	<tr height="30px">
		<td style="padding-left: 0px; text-align: left; font-size: 1.1em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Tgl Ijin Prinsip
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.1em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; font-size: 1.1em; vertical-align: top">
			<FONT face="Arial" color="#000">
				<!--input style="font-family: Arial; font-size: 1.1em" type="text" id="datepicker-ex30" name="barupajaktglbayar" class="textboxtgl"-->
<%
			response.write(tglxx)
%>
			</FONT>
		</td>
	</tr>
	<tr height="10px">
		<td style="padding-left: 0px; text-align: left; font-size: 1.1em; vertical-align: top">									
		</td>
	</tr>

</table>	

<table bordercolor="#999999" style="text-align: left; margin-left: 50px; margin-right: auto; margin-top: auto;" border="1" cellspacing="0" cellpadding="0">
	<tr height="40px" style="background-color: #DEDEDE">
		<td style="padding-left: auto; text-align: center; font-size: 1.1em;" width="50px">
			<FONT face="Arial" color="#000"><b>NO</b></FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.1em;" width="100px">
			<FONT face="Arial" color="#000"><b>TGL</b></FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.1em;" width="310px">
			<FONT face="Arial" color="#000"><b>PROSES</b></FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.1em;" width="500px">
			<FONT face="Arial" color="#000"><b>KETERANGAN</b></FONT>
		</td>
		<td colspan="6" style="padding-left: auto; text-align: center; font-size: 1.1em;" width="300px">
			<FONT face="Arial" color="#000"><b>DOKUMEN DIUNGGAH</b></FONT>
		</td>
	</tr>
<%
dd = 0
'================================================================================================================
' PERMOHONAN (AWAL)
'================================================================================================================

If jmliji=1 Then
	dd = dd + 1
	folderx = ""
	filesuratx = ""
	fileipx = ""
	filetorx = ""
	filerabx = ""
	SQL3 = "SELECT * FROM pengadaan WHERE ID=" & adaid
'response.write(SQL3)
	myCmd.CommandText = SQL3
	myConn.Open()
	myReader = myCmd.ExecuteReader()
	If myReader.HasRows Then
		While myReader.Read()
			idx = myReader("id").ToString()
			nosuratx = myReader("kode").ToString()
			tglx = myReader("tglmohon").ToString()
			nilaix = myReader("nilai").ToString()
			nilaihpsx = myReader("nilaihps").ToString()
			nilaijml = Val(nilaix)
			nilaihps = Val(nilaihpsx)
			cuserx = myReader("createduser").ToString()
			nostepx = myReader("nostep").ToString()
			nostepv = Val(nostepx)
			statusrevisix = myReader("statusrevisi").ToString()
			statusrejectx = Val(myReader("statusreject").ToString())
			folderx = myReader("namafolder").ToString()
			filesuratx = myReader("namafilesurat").ToString()
			fileipx = myReader("namafileip").ToString()
			filetorx = myReader("namafiletor").ToString()
			filerabx = myReader("namafilerab").ToString()
			stepstr = nostepx
			dtgl = Convert.ToDateTime(tglx)
			tglxx = dtgl.ToString("dd-MM-yyyy")
			tglawalx = tglx
			'Dim oDate As DateTime = Convert.ToDateTime(tglx)
			jmliji = 1

		End While
	End If
	myConn.Close()


	If filesuratx.length() > 2 Then
		jenisarr = filesuratx.Split(".")
		jeniscount = jenisarr.count()
		jenisfile1 = jenisarr(jeniscount-1)
	End If
	If fileipx.length() > 2 Then
		jenisarr = fileipx.Split(".")
		jeniscount = jenisarr.count()
		jenisfile2 = jenisarr(jeniscount-1)
	End If
	If filetorx.length() > 2 Then
		jenisarr = filetorx.Split(".")
		jeniscount = jenisarr.count()
		jenisfile3 = jenisarr(jeniscount-1)
	End If
	If filerabx.length() > 2 Then
		jenisarr = filerabx.Split(".")
		jeniscount = jenisarr.count()
		jenisfile4 = jenisarr(jeniscount-1)
	End If


%>
	<tr height="40px" style="background-color: #FFFFFF">
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="50px">
			<FONT face="Arial" color="#000">
			<a href="#" onclick="viewijin(<% response.write(userid) %>,<% response.write(adaid) %>);">
			<% response.write(dd) %>
			</a>
			</FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="150px">
			<FONT face="Arial" color="#000">
			<a href="#" onclick="viewijin(<% response.write(userid) %>,<% response.write(adaid) %>);">
			<% response.write(tglxx) %>
			</a>
			</FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="220px">
			<FONT face="Arial" color="#000">
			<a href="#" onclick="viewijin(<% response.write(userid) %>,<% response.write(adaid) %>);">
			Permohonan/Ijin Prinsip
			</a>
			</FONT>
		</td>
		<td style="padding-left: 10px; text-align: left; font-size: 1.0em;" width="400px">
			<FONT face="Arial" color="#000">
<%
				response.write("<i>Anggaran</i>: Rp. " & nilaijml.ToString("#,###"))
%>		
			</FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="50px">
<%
		If filesuratx.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & filesuratx) %>" target="_blank">
<%
				filestrx = "." & folderx & filesuratx
				If jenisfile1="pdf" Then
					filestrx = "./images/pdficon1.png"
				End If
%>
			<img id="uploadPreview1" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview1" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="50px">
<%
		If fileipx.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & fileipx) %>" target="_blank">
<%
				filestrx = "." & folderx & fileipx
				If jenisfile2="pdf" Then
					filestrx = "./images/pdficon1.png"
				End If
%>
			<img id="uploadPreview2" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview2" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="50px">
<%
		If filetorx.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & filetorx) %>" target="_blank">
<%
				filestrx = "." & folderx & filetorx
				If jenisfile3="pdf" Then
					filestrx = "./images/pdficon1.png"
				End If
%>
			<img id="uploadPreview3" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview2" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="50px">
<%
		If filerabx.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & filerabx) %>" target="_blank">
<%
				filestrx = "." & folderx & filerabx
				If jenisfile4="pdf" Then
					filestrx = "./images/pdficon1.png"
				End If
%>
			<img id="uploadPreview4" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview2" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td colspan="2" style="padding-left: auto; text-align: center; font-size: 1.0em;">
		</td>
	</tr>
<%
End If

If jmlusu=1 Then
	dd = dd + 1
'==============================================================================
' 2. BACA DATA PROSES USULAN
'==============================================================================
	SQL3 = "SELECT * FROM pengadaan_prosesusulan WHERE pengadaanID=" & adaid
	myCmd.CommandText = SQL3
	myConn.Open()
	myReader = myCmd.ExecuteReader()
	file1x = ""
	file2x = ""
	If myReader.HasRows Then
		While myReader.Read()
			idx = myReader("ID").ToString()
			tglx = myReader("tgl").ToString()
			hpsx = Val(myReader("nilaihps").ToString())
			folderx = myReader("namafolder").ToString()
			file1x = myReader("namafile1").ToString()
			file2x = myReader("namafile2").ToString()
			file3x = myReader("namafile3").ToString()
			file4x = myReader("namafile4").ToString()
			file5x = myReader("namafile5").ToString()
			file6x = myReader("namafile6").ToString()
			ketx = myReader("keterangan").ToString()
			dtgl = Convert.ToDateTime(tglx)
			tglxx = dtgl.ToString("dd-MM-yyyy")
			tglakhirx = tglx
			'Dim oDate As DateTime = Convert.ToDateTime(tglx)
		End While
	End If
	myConn.Close()
	
	If file1x.length() > 2 Then
		jenisarr = file1x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile1 = jenisarr(jeniscount-1)
	End If
	If file2x.length() > 2 Then
		jenisarr = file2x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile2 = jenisarr(jeniscount-1)
	End If
	If file3x.length() > 2 Then
		jenisarr = file3x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile3 = jenisarr(jeniscount-1)
	End If
	If file4x.length() > 2 Then
		jenisarr = file4x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile4 = jenisarr(jeniscount-1)
	End If
	If file5x.length() > 2 Then
		jenisarr = file5x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile3 = jenisarr(jeniscount-1)
	End If
	If file6x.length() > 2 Then
		jenisarr = file6x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile4 = jenisarr(jeniscount-1)
	End If
	
%>
	<tr height="40px" style="background-color: #FFFFFF">
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;">
			<FONT face="Arial" color="#000">
			<a href="#" onclick="viewusul(<% response.write(userid) %>,<% response.write(adaid) %>);">
			<% response.write(dd) %>
			</a>
			</FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;">
			<FONT face="Arial" color="#000">
			<a href="#" onclick="viewusul(<% response.write(userid) %>,<% response.write(adaid) %>);">
			<% response.write(tglxx) %>
			</a>
			</FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;">
			<FONT face="Arial" color="#000">
			<a href="#" onclick="viewusul(<% response.write(userid) %>,<% response.write(adaid) %>);">
			Proses Usulan
			</a>
			</FONT>
		</td>
		<td style="padding-left: 10px; text-align: left; font-size: 1.0em;">
			<FONT face="Arial" color="#000">
<%
				response.write("<i>HPS</i>: Rp. " & hpsx.ToString("#,###"))
%>
			</FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="50px">
<%
		If file1x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file1x) %>" target="_blank">
<%
			filestrx = "." & folderx & file1x
			If jenisfile1="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview1" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview1" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="50px">
<%
		If file2x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file2x) %>" target="_blank">
<%
			filestrx = "." & folderx & file2x
			If jenisfile2="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview2" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview2" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="50px">
<%
		If file3x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file3x) %>" target="_blank">
<%
			filestrx = "." & folderx & file3x
			If jenisfile3="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview2" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview2" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="50px">
<%
		If file4x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file4x) %>" target="_blank">
<%
			filestrx = "." & folderx & file4x
			If jenisfile4="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview2" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview2" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="50px">
<%
		If file5x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file5x) %>" target="_blank">
<%
			filestrx = "." & folderx & file5x
			If jenisfile5="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview2" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview2" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="50px">
<%
		If file6x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file6x) %>" target="_blank">
<%
			filestrx = "." & folderx & file6x
			If jenisfile6="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview2" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview2" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
	</tr>
<%
End If

If jmlaan=1 Then
	dd = dd + 1
'==============================================================================
' 3. BACA DATA AANWIJZING
'==============================================================================
	SQL3 = "SELECT * FROM pengadaan_aanwijzing WHERE pengadaanID=" & adaid
	myCmd.CommandText = SQL3
	myConn.Open()
	myReader = myCmd.ExecuteReader()

	If myReader.HasRows Then
		While myReader.Read()
			idx = myReader("ID").ToString()
			tglx = myReader("tgl").ToString()
			folderx = myReader("namafolder").ToString()
			file1x = myReader("namafile1").ToString()
			file2x = myReader("namafile2").ToString()
			ketx = myReader("keterangan").ToString()
			dtgl = Convert.ToDateTime(tglx)
			tglxx = dtgl.ToString("dd-MM-yyyy")
			tglakhirx = tglx
			'Dim oDate As DateTime = Convert.ToDateTime(tglx)
		End While
	End If
	myConn.Close()

	
	If file1x.length() > 2 Then
		jenisarr = file1x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile1 = jenisarr(jeniscount-1)
	End If
	If file2x.length() > 2 Then
		jenisarr = file2x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile2 = jenisarr(jeniscount-1)
	End If
%>
	<tr height="40px" style="background-color: #FFFFFF">
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewaanwij(<% response.write(userid) %>,<% response.write(adaid) %>);">
				<% response.write(dd) %>
				</a>
			</FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewaanwij(<% response.write(userid) %>,<% response.write(adaid) %>);">
				<% response.write(tglxx) %>
				</a>
			</FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;">
			<FONT face="Arial" color="#000">
			<a href="#" onclick="viewaanwij(<% response.write(userid) %>,<% response.write(adaid) %>);">
			Aanwijzing
			</a>
			</FONT>
		</td>
		<td style="padding-left: 10px; text-align: left; font-size: 1.0em;">
			<FONT face="Arial" color="#000">		
			</FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="50px">
<%
		If file1x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file1x) %>" target="_blank">
<%
			filestrx = "." & folderx & file1x
			If jenisfile1="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview1" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview1" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="50px">
<%
		If file2x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file2x) %>" target="_blank">
<%
			filestrx = "." & folderx & file2x
			If jenisfile2="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview2" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview2" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td colspan="4" style="padding-left: auto; text-align: center; font-size: 1.0em;">
		</td>

	</tr>
<%
End If

If jmlpoc=1 Then
	dd = dd + 1
'==============================================================================
' 4. BACA DATA POC
'==============================================================================
	SQL3 = "SELECT * FROM pengadaan_poc WHERE pengadaanID=" & adaid
	myCmd.CommandText = SQL3
	myConn.Open()
	myReader = myCmd.ExecuteReader()

	If myReader.HasRows Then
		While myReader.Read()
			idx = myReader("ID").ToString()
			tglx = myReader("tgl").ToString()
			folderx = myReader("namafolder").ToString()
			file1x = myReader("namafile1").ToString()
			file2x = myReader("namafile2").ToString()
			ketx = myReader("keterangan").ToString()
			dtgl = Convert.ToDateTime(tglx)
			tglxx = dtgl.ToString("dd-MM-yyyy")
			tglakhirx = tglx
			'Dim oDate As DateTime = Convert.ToDateTime(tglx)
		End While
	End If
	myConn.Close()
	
	If file1x.length() > 2 Then
		jenisarr = file1x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile1 = jenisarr(jeniscount-1)
	End If
	If file2x.length() > 2 Then
		jenisarr = file2x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile2 = jenisarr(jeniscount-1)
	End If
%>
	<tr height="40px" style="background-color: #FFFFFF">
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewpoc(<% response.write(userid) %>,<% response.write(adaid) %>);">
				<% response.write(dd) %>
				</a>
			</FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewpoc(<% response.write(userid) %>,<% response.write(adaid) %>);">
				<% response.write(tglxx) %>
				</a>
			</FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;">
			<FONT face="Arial" color="#000">
			<a href="#" onclick="viewpoc(<% response.write(userid) %>,<% response.write(adaid) %>);">
			Proof of Concept (PoC)
			</a>
			</FONT>
		</td>
		<td style="padding-left: 10px; text-align: left; font-size: 1.0em;">
			<FONT face="Arial" color="#000">		
			</FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="50px">
<%
		If file1x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file1x) %>" target="_blank">
<%
			filestrx = "." & folderx & file1x
			If jenisfile1="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview1" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview1" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="50px">
<%
		If file2x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file2x) %>" target="_blank">
<%
			filestrx = "." & folderx & file2x
			If jenisfile2="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview2" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview2" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td colspan="4" style="padding-left: auto; text-align: center; font-size: 1.0em;">
		</td>

	</tr>
<%
End If

If jmleva=1 Then
	dd = dd + 1
'==============================================================================
' 5. BACA DATA EVALUASI ADMIN & TEKNIS
'==============================================================================
	SQL3 = "SELECT * FROM pengadaan_evaluasi WHERE pengadaanID=" & adaid
	myCmd.CommandText = SQL3
	myConn.Open()
	myReader = myCmd.ExecuteReader()

	If myReader.HasRows Then
		While myReader.Read()
			idx = myReader("ID").ToString()
			tglx = myReader("tgl").ToString()
			folderx = myReader("namafolder").ToString()
			file1x = myReader("namafile1").ToString()
			file2x = myReader("namafile2").ToString()
			ketx = myReader("keterangan").ToString()
			dtgl = Convert.ToDateTime(tglx)
			tglxx = dtgl.ToString("dd-MM-yyyy")
			tglakhirx = tglx
			'Dim oDate As DateTime = Convert.ToDateTime(tglx)
		End While
	End If
	myConn.Close()
	
	If file1x.length() > 2 Then
		jenisarr = file1x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile1 = jenisarr(jeniscount-1)
	End If
	If file2x.length() > 2 Then
		jenisarr = file2x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile2 = jenisarr(jeniscount-1)
	End If
%>
	<tr height="40px" style="background-color: #FFFFFF">
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="vieweval(<% response.write(userid) %>,<% response.write(adaid) %>);">
				<% response.write(dd) %>
				</a>
			</FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="vieweval(<% response.write(userid) %>,<% response.write(adaid) %>);">
				<% response.write(tglxx) %>
				</a>
			</FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;">
			<FONT face="Arial" color="#000">
			<a href="#" onclick="vieweval(<% response.write(userid) %>,<% response.write(adaid) %>);">
			Evaluasi Adm & Teknis
			</a>
			</FONT>
		</td>
		<td style="padding-left: 10px; text-align: left; font-size: 1.0em;">
			<FONT face="Arial" color="#000">		
			</FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="50px">
<%
		If file1x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file1x) %>" target="_blank">
<%
			filestrx = "." & folderx & file1x
			If jenisfile1="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview1" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview1" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="50px">
<%
		If file2x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file2x) %>" target="_blank">
<%
			filestrx = "." & folderx & file2x
			If jenisfile2="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview2" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview2" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td colspan="4" style="padding-left: auto; text-align: center; font-size: 1.0em;">
		</td>

	</tr>
<%
End If

If jmlneg=1 Then
	dd = dd + 1
'==============================================================================
' 6. BACA DATA NEGOSIASI
'==============================================================================

	
	SQL3 = "SELECT * FROM pengadaan_negosiasi WHERE pengadaanID=" & adaid
	myCmd.CommandText = SQL3
	myConn.Open()
	myReader = myCmd.ExecuteReader()
	d = 1
	If myReader.HasRows Then
		While myReader.Read()
			idx = myReader("ID").ToString()
			tglx = myReader("tgl").ToString()
			jmlprenegox = Val(myReader("jmlprenego").ToString())
			jmlpostnegox = Val(myReader("jmlpostnego").ToString())
			jmlpostnegox2 = Val(myReader("jmlpostnego2").ToString())
			jmlpostnegox3 = Val(myReader("jmlpostnego3").ToString())
			jmlpostnegox4 = Val(myReader("jmlpostnego4").ToString())
			jmlpostnegox5 = Val(myReader("jmlpostnego5").ToString())
			jmlpostnegox6 = Val(myReader("jmlpostnego6").ToString())
			vendorx = myReader("vendor").ToString()
			vendorx2 = myReader("vendor2").ToString()
			vendorx3 = myReader("vendor3").ToString()
			vendorx4 = myReader("vendor4").ToString()
			vendorx5 = myReader("vendor5").ToString()
			vendorx6 = myReader("vendor6").ToString()
			folderx = myReader("namafolder").ToString()
			file1x = myReader("namafile1").ToString()
			file2x = myReader("namafile2").ToString()
			ketx = myReader("keterangan").ToString()
			dtgl = Convert.ToDateTime(tglx)
			tglxx = dtgl.ToString("dd-MM-yyyy")
			tglakhirx = tglx
			vendorz(1) = vendorx
			jmlnegoz(1) = jmlpostnegox
			negostr = vendorx & " (Rp. " & jmlpostnegox.ToString("#,###") & ")"
			If vendorx2.length() > 2 Then
				d = d + 1
				vendorz(d) = vendorx2
				jmlnegoz(d) = jmlpostnegox2
				negostr = negostr & "<br>" & vendorx2 & " (Rp. " & jmlpostnegox2.ToString("#,###") & ")"
			End If
			If vendorx3.length() > 2 Then
				d = d + 1
				vendorz(d) = vendorx3
				jmlnegoz(d) = jmlpostnegox3
				negostr = negostr & "<br>" & vendorx3 & " (Rp. " & jmlpostnegox3.ToString("#,###") & ")"
			End If
			If vendorx4.length() > 2 Then
				d = d + 1
				vendorz(d) = vendorx4
				jmlnegoz(d) = jmlpostnegox4
				negostr = negostr & "<br>" & vendorx4 & " (Rp. " & jmlpostnegox4.ToString("#,###") & ")"
			End If
			If vendorx5.length() > 2 Then
				d = d + 1
				vendorz(d) = vendorx5
				jmlnegoz(d) = jmlpostnegox5
				negostr = negostr & "<br>" & vendorx5 & " (Rp. " & jmlpostnegox5.ToString("#,###") & ")"
			End If
			If vendorx6.length() > 2 Then
				d = d + 1
				vendorz(d) = vendorx6
				jmlnegoz(d) = jmlpostnegox6
				negostr = negostr & "<br>" & vendorx6 & " (Rp. " & jmlpostnegox6.ToString("#,###") & ")"
			End If
			'Dim oDate As DateTime = Convert.ToDateTime(tglx)
		End While
	End If
	myConn.Close()


	If file1x.length() > 2 Then
		jenisarr = file1x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile1 = jenisarr(jeniscount-1)
	End If
	If file2x.length() > 2 Then
		jenisarr = file2x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile2 = jenisarr(jeniscount-1)
	End If
%>
	<tr height="40px" style="background-color: #FFFFFF">
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewnego(<% response.write(userid) %>,<% response.write(adaid) %>);">
				<% response.write(dd) %>
				</a>
			</FONT>
		</td>
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewnego(<% response.write(userid) %>,<% response.write(adaid) %>);">
				<% response.write(tglxx) %>
				</a>
			</FONT>
		</td>
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;">
			<FONT face="Arial" color="#000">
			<a href="#" onclick="viewnego(<% response.write(userid) %>,<% response.write(adaid) %>);">
			Negosiasi
			</a>
			</FONT>
		</td>
		<td style="padding-top: 8px; padding-bottom: 8px; padding-left: 10px; text-align: left; font-size: 1.0em; vertical-align: top;">
			<FONT face="Arial" color="#000">
<%
				'response.write("Hasil Nego: Rp. " & jmlpostnegox.ToString("#,###") & "<br>")
				'response.write("Vendor: " & vendorx)
				response.write(negostr)
%>			
			</FONT>
		</td>
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;" width="50px">
<%
		If file1x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file1x) %>" target="_blank">
<%
			filestrx = "." & folderx & file1x
			If jenisfile1="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview1" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview1" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;" width="50px">
<%
		If file2x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file2x) %>" target="_blank">
<%
			filestrx = "." & folderx & file2x
			If jenisfile2="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview2" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview2" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td colspan="4" style="padding-left: auto; text-align: center; font-size: 1.0em;">
		</td>
	</tr>
<%
End If

If jmlnkp=1 Then
	dd = dd + 1
'==============================================================================
' 7. BACA DATA NKP
'==============================================================================
	SQL3 = "SELECT * FROM pengadaan_nkp WHERE pengadaanID=" & adaid
	myCmd.CommandText = SQL3
	myConn.Open()
	myReader = myCmd.ExecuteReader()

	If myReader.HasRows Then
		While myReader.Read()
			idx = myReader("ID").ToString()
			tglx = myReader("tgl").ToString()
			nomory = myReader("nomor").ToString()
			folderx = myReader("namafolder").ToString()
			file1x = myReader("namafile1").ToString()
			file2x = myReader("namafile2").ToString()
			ketx = myReader("keterangan").ToString()
			dtgl = Convert.ToDateTime(tglx)
			tglxx = dtgl.ToString("dd-MM-yyyy")
			tglakhirx = tglx
			'Dim oDate As DateTime = Convert.ToDateTime(tglx)
		End While
	End If
	myConn.Close()
	
	If file1x.length() > 2 Then
		jenisarr = file1x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile1 = jenisarr(jeniscount-1)
	End If
	If file2x.length() > 2 Then
		jenisarr = file2x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile2 = jenisarr(jeniscount-1)
	End If
%>
	<tr height="40px" style="background-color: #FFFFFF">
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewnkp(<% response.write(userid) %>,<% response.write(adaid) %>);">
				<% response.write(dd) %>
				</a>
			</FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewnkp(<% response.write(userid) %>,<% response.write(adaid) %>);">
				<% response.write(tglxx) %>
				</a>
			</FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewnkp(<% response.write(userid) %>,<% response.write(adaid) %>);">
				Nota Keputusan Pengadaan (NKP)
				</a>
			</FONT>
		</td>
		<td style="padding-left: 10px; text-align: left; font-size: 1.0em;">
			<FONT face="Arial" color="#000">
<%
			response.write(nomory)
%>
			</FONT>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="50px">
<%
		If file1x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file1x) %>" target="_blank">
<%
			filestrx = "." & folderx & file1x
			If jenisfile1="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview1" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview1" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="50px">
<%
		If file2x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file2x) %>" target="_blank">
<%
			filestrx = "." & folderx & file2x
			If jenisfile2="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview2" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview2" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td colspan="4" style="padding-left: auto; text-align: center; font-size: 1.0em;">
		</td>

	</tr>
<%
End If

If jmlsan=1 Then
	dd = dd + 1
'==============================================================================
' 8. BACA DATA SANGGAHAN
'==============================================================================
	SQL3 = "SELECT * FROM pengadaan_sanggah WHERE pengadaanID=" & adaid
	myCmd.CommandText = SQL3
	myConn.Open()
	myReader = myCmd.ExecuteReader()

	If myReader.HasRows Then
		While myReader.Read()
			idx = myReader("ID").ToString()
			tglx = myReader("tgl").ToString()
			nomorz(1) = myReader("nomor").ToString()
			nomorz(2) = myReader("nomor2").ToString()
			nomorz(3) = myReader("nomor3").ToString()
			nomorz(4) = myReader("nomor4").ToString()
			nomorz(5) = myReader("nomor5").ToString()
			nomorz(6) = myReader("nomor6").ToString()
			folderx = myReader("namafolder").ToString()
			file1x = myReader("namafile1").ToString()
			file2x = myReader("namafile2").ToString()
			ketx = myReader("keterangan").ToString()
			dtgl = Convert.ToDateTime(tglx)
			tglxx = dtgl.ToString("dd-MM-yyyy")
			tglakhirx = tglx
			'Dim oDate As DateTime = Convert.ToDateTime(tglx)
		End While
	End If
	myConn.Close()
	
	For c = 1 to 6
		If nomorz(c).length() > 3 Then
			d = d + 1
			'vendorzz(d) = vendorz(c)
			'jmlnegozz(d) = jmlnegoz(c)
			'vendorstr = vendorstr & vendorz(c) & " (Rp. " & jmlnegoz(c).ToString("#,###") & ") - " & nomorz(c) & "<br>"
			vendorstr = vendorstr & nomorz(c) & "<br>"
		End If
	Next c
	
	If file1x.length() > 2 Then
		jenisarr = file1x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile1 = jenisarr(jeniscount-1)
	End If
	If file2x.length() > 2 Then
		jenisarr = file2x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile2 = jenisarr(jeniscount-1)
	End If
%>
	<tr height="40px" style="background-color: #FFFFFF">
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewsanggah(<% response.write(userid) %>,<% response.write(adaid) %>);">
				<% response.write(dd) %>
				</a>
			</FONT>
		</td>
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewsanggah(<% response.write(userid) %>,<% response.write(adaid) %>);">
				<% response.write(tglxx) %>
				</a>
			</FONT>
		</td>
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewsanggah(<% response.write(userid) %>,<% response.write(adaid) %>);">
				Sanggahan
				</a>
			</FONT>
		</td>
		<td style="padding-top: 8px; padding-bottom: 8px; padding-left: 10px; text-align: left; font-size: 1.0em; vertical-align: top;">
			<FONT face="Arial" color="#000">
<%
			response.write(vendorstr)
%>			
			</FONT>
		</td>
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;" width="50px">
<%
		If file1x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file1x) %>" target="_blank">
<%
			filestrx = "." & folderx & file1x
			If jenisfile1="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview1" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview1" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;" width="50px">
<%
		If file2x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file2x) %>" target="_blank">
<%
			filestrx = "." & folderx & file2x
			If jenisfile2="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview2" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview2" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td colspan="4" style="padding-left: auto; text-align: center; font-size: 1.0em;">
		</td>

	</tr>
<%
End If

If jmlpen=1 Then
	dd = dd + 1
'==============================================================================
' 9. BACA DATA PENUNJUKAN PEMENANG
'==============================================================================
	'Dim vendorx as String = ""
	'Dim vendorx2 as String = ""
	'Dim vendorx3 as String = ""
	'Dim vendorx4 as String = ""
	'Dim vendorx5 as String = ""
	'Dim vendorx6 as String = ""
	'Dim jmlprenegox as Long = 0
	'Dim jmlpostnegox as Long = 0
	'Dim jmlpostnegox2 as Long = 0
	'Dim jmlpostnegox3 as Long = 0
	'Dim jmlpostnegox4 as Long = 0
	'Dim jmlpostnegox5 as Long = 0
	'Dim jmlpostnegox6 as Long = 0

	SQL3 = "SELECT * FROM pengadaan_negosiasi WHERE pengadaanID=" & adaid
	myCmd.CommandText = SQL3
	myConn.Open()
	myReader = myCmd.ExecuteReader()

	If myReader.HasRows Then
		While myReader.Read()
			jmlprenegox = Val(myReader("jmlprenego").ToString())
			jmlpostnegox = Val(myReader("jmlpostnego").ToString())
			jmlpostnegox2 = Val(myReader("jmlpostnego2").ToString())
			jmlpostnegox3 = Val(myReader("jmlpostnego3").ToString())
			jmlpostnegox4 = Val(myReader("jmlpostnego4").ToString())
			jmlpostnegox5 = Val(myReader("jmlpostnego5").ToString())
			jmlpostnegox6 = Val(myReader("jmlpostnego6").ToString())
			vendorx = myReader("vendor").ToString()
			vendorx2 = myReader("vendor2").ToString()
			vendorx3 = myReader("vendor3").ToString()
			vendorx4 = myReader("vendor4").ToString()
			vendorx5 = myReader("vendor5").ToString()
			vendorx6 = myReader("vendor6").ToString()
		End While
	End If
	myConn.Close()
	
	SQL3 = "SELECT * FROM pengadaan_penunjukan WHERE pengadaanID=" & adaid
	myCmd.CommandText = SQL3
	myConn.Open()
	myReader = myCmd.ExecuteReader()

	If myReader.HasRows Then
		While myReader.Read()
			idx = myReader("ID").ToString()
			tglx = myReader("tgl").ToString()
			vendorz(1) = myReader("vendor").ToString()
			vendorz(2) = myReader("vendor2").ToString()
			vendorz(3) = myReader("vendor3").ToString()
			vendorz(4) = myReader("vendor4").ToString()
			vendorz(5) = myReader("vendor5").ToString()
			vendorz(6) = myReader("vendor6").ToString()
			nomorz(1) = myReader("nomor").ToString()
			nomorz(2) = myReader("nomor2").ToString()
			nomorz(3) = myReader("nomor3").ToString()
			nomorz(4) = myReader("nomor4").ToString()
			nomorz(5) = myReader("nomor5").ToString()
			nomorz(6) = myReader("nomor6").ToString()
			folderx = myReader("namafolder").ToString()
			file1x = myReader("namafile1").ToString()
			file2x = myReader("namafile2").ToString()
			ketx = myReader("keterangan").ToString()
			dtgl = Convert.ToDateTime(tglx)
			tglxx = dtgl.ToString("dd-MM-yyyy")
			tglakhirx = tglx
			'Dim oDate As DateTime = Convert.ToDateTime(tglx)
		End While
	End If
	myConn.Close()

'------------------------------------------------------
' Baca jmlnego dari tabel pengadaan_negosiasi	
'------------------------------------------------------
	d = 0
	vendorstr = ""
	For c = 1 to 6
		If vendorz(c).length() > 3 Then
			d = d + 1
			vendorzz(d) = vendorz(c)
			If vendorz(c) = vendorx Then
				jmlnegoz(c) = jmlpostnegox
			Else If vendorz(c) = vendorx2 Then
				jmlnegoz(c) = jmlpostnegox2
			Else If vendorz(c) = vendorx3 Then
				jmlnegoz(c) = jmlpostnegox3
			Else If vendorz(c) = vendorx4 Then
				jmlnegoz(c) = jmlpostnegox4
			Else If vendorz(c) = vendorx5 Then
				jmlnegoz(c) = jmlpostnegox5
			Else If vendorz(c) = vendorx6 Then
				jmlnegoz(c) = jmlpostnegox6
			End If
			jmlnegozz(d) = jmlnegoz(c)
			vendorstr = vendorstr & vendorz(c) & " (Rp. " & jmlnegoz(c).ToString("#,###") & ") - " & nomorz(c) & "<br>"
		End If
	Next c
	
	
	If file1x.length() > 2 Then
		jenisarr = file1x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile1 = jenisarr(jeniscount-1)
	End If
	If file2x.length() > 2 Then
		jenisarr = file2x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile2 = jenisarr(jeniscount-1)
	End If
%>
	<tr height="40px" style="background-color: #FFFFFF">
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewpenunjukan(<% response.write(userid) %>,<% response.write(adaid) %>);">
				<% response.write(dd) %>
				</a>
			</FONT>
		</td>
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewpenunjukan(<% response.write(userid) %>,<% response.write(adaid) %>);">
				<% response.write(tglxx) %>
				</a>
			</FONT>
		</td>
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewpenunjukan(<% response.write(userid) %>,<% response.write(adaid) %>);">
				Penunjukan Pemenang
				</a>
			</FONT>
		</td>
		<td style="padding-top: 8px; padding-bottom: 8px; padding-left: 10px; text-align: left; font-size: 1.0em; vertical-align: top;">
			<FONT face="Arial" color="#000">
<%
				response.write(vendorstr)
%>			
			</FONT>
		</td>
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;" width="50px">
<%
		If file1x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file1x) %>" target="_blank">
<%
			filestrx = "." & folderx & file1x
			If jenisfile1="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview1" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview1" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;" width="50px">
<%
		If file2x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file2x) %>" target="_blank">
<%
			filestrx = "." & folderx & file2x
			If jenisfile2="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview2" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview2" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td colspan="4" style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;">
		</td>

	</tr>
<%
End If

If jmlspk=1 Then
	dd = dd + 1
'==============================================================================
' 10. BACA DATA SPK
'==============================================================================
	SQL3 = "SELECT * FROM pengadaan_spk WHERE pengadaanID=" & adaid
	myCmd.CommandText = SQL3
	myConn.Open()
	myReader = myCmd.ExecuteReader()

	If myReader.HasRows Then
		While myReader.Read()
			idx = myReader("ID").ToString()
			tglx = myReader("tgl").ToString()
			nomorz(1) = myReader("nomor").ToString()
			nomorz(2) = myReader("nomor2").ToString()
			nomorz(3) = myReader("nomor3").ToString()
			nomorz(4) = myReader("nomor4").ToString()
			nomorz(5) = myReader("nomor5").ToString()
			nomorz(6) = myReader("nomor6").ToString()
			folderx = myReader("namafolder").ToString()
			file1x = myReader("namafile1").ToString()
			file2x = myReader("namafile2").ToString()
			file3x = myReader("namafile3").ToString()
			dtgl = Convert.ToDateTime(tglx)
			tglxx = dtgl.ToString("dd-MM-yyyy")
			'tglspkx = tglx
			tglakhirx = tglx
			'Dim oDate As DateTime = Convert.ToDateTime(tglx)
		End While
	End If
	myConn.Close()
	
	vendorstr = ""
	For c = 1 to d
		vendorstr = vendorstr & vendorzz(c) & " (Rp. " & jmlnegozz(c).ToString("#,###") & ") - " & nomorz(c) & "<br>"
	Next c
	
	If file1x.length() > 2 Then
		jenisarr = file1x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile1 = jenisarr(jeniscount-1)
	End If
	If file2x.length() > 2 Then
		jenisarr = file2x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile2 = jenisarr(jeniscount-1)
	End If
%>
	<tr height="40px" style="background-color: #FFFFFF">
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewspk(<% response.write(userid) %>,<% response.write(adaid) %>);">
				<% response.write(dd) %>
				</a>
			</FONT>
		</td>
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewspk(<% response.write(userid) %>,<% response.write(adaid) %>);">
				<% response.write(tglxx) %>
				</a>
			</FONT>
		</td>
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewspk(<% response.write(userid) %>,<% response.write(adaid) %>);">
				SPK
				</a>
			</FONT>
		</td>
		<td style="padding-top: 8px; padding-bottom: 8px; padding-left: 10px; text-align: left; font-size: 1.0em; vertical-align: top;">
			<FONT face="Arial" color="#000">
<%
				response.write(vendorstr)
%>		
			</FONT>
		</td>
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;" width="50px">
<%
		If file1x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file1x) %>" target="_blank">
<%
			filestrx = "." & folderx & file1x
			If jenisfile1="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview1" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview1" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;" width="50px">
<%
		If file2x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file2x) %>" target="_blank">
<%
			filestrx = "." & folderx & file2x
			If jenisfile2="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview2" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview2" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td colspan="4" style="padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;">
		</td>

	</tr>
<%
End If


If jmlpks=1 Then
	dd = dd + 1
'==============================================================================
' 11. BACA DATA PKS
'==============================================================================
	SQL3 = "SELECT * FROM pengadaan_pks WHERE pengadaanID=" & adaid
	myCmd.CommandText = SQL3
	myConn.Open()
	myReader = myCmd.ExecuteReader()

	If myReader.HasRows Then
		While myReader.Read()
			idx = myReader("ID").ToString()
			tglx = myReader("tgl").ToString()
			nomorz(1) = myReader("nomor").ToString()
			nomorz(2) = myReader("nomor2").ToString()
			nomorz(3) = myReader("nomor3").ToString()
			nomorz(4) = myReader("nomor4").ToString()
			nomorz(5) = myReader("nomor5").ToString()
			nomorz(6) = myReader("nomor6").ToString()
			folderx = myReader("namafolder").ToString()
			file1x = myReader("namafile1").ToString()
			file2x = myReader("namafile2").ToString()
			'ketx = myReader("keterangan").ToString()
			dtgl = Convert.ToDateTime(tglx)
			tglxx = dtgl.ToString("dd-MM-yyyy")
			'Dim oDate As DateTime = Convert.ToDateTime(tglx)
			'tglakhirx = tglx
		End While
	End If
	myConn.Close()
	
	vendorstr = ""
	For c = 1 to d
		vendorstr = vendorstr & vendorzz(c) & " (Rp. " & jmlnegozz(c).ToString("#,###") & ") - " & nomorz(c) & "<br>"
	Next c	
	
	If file1x.length() > 2 Then
		jenisarr = file1x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile1 = jenisarr(jeniscount-1)
	End If
	If file2x.length() > 2 Then
		jenisarr = file2x.Split(".")
		jeniscount = jenisarr.count()
		jenisfile2 = jenisarr(jeniscount-1)
	End If

%>
	<tr height="40px" style="background-color: #FFFFFF">
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewpks(<% response.write(userid) %>,<% response.write(adaid) %>);">
				<% response.write(dd) %>
				</a>
			</FONT>
		</td>
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewpks(<% response.write(userid) %>,<% response.write(adaid) %>);">
				<% response.write(tglxx) %>
				</a>
			</FONT>
		</td>
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;">
			<FONT face="Arial" color="#000">
				<a href="#" onclick="viewpks(<% response.write(userid) %>,<% response.write(adaid) %>);">
				PKS
				</a>
			</FONT>
		</td>
		<td style="padding-top: 8px; padding-left: 10px; padding-bottom: 8px; text-align: left; font-size: 1.0em; vertical-align: top;">
			<FONT face="Arial" color="#000">
<%
				response.write(vendorstr)
%>		
			</FONT>
		</td>
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;" width="50px">
<%
		If file1x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file1x) %>" target="_blank">
<%
			filestrx = "." & folderx & file1x
			If jenisfile1="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview1" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview1" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;" width="50px">
<%
		If file2x.length() > 2 Then
%>
			<a id="linkimage" href="<% response.write("." & folderx & file2x) %>" target="_blank">
<%
			filestrx = "." & folderx & file2x
			If jenisfile2="pdf" Then
				filestrx = "./images/pdficon1.png"
			End If
%>
			<img id="uploadPreview2" src="<% response.write(filestrx) %>" height="30px" width="25px"/><br/>
			</a>
<%
		Else
%>
			<a id="linkimage" href="<% response.write("./images/nofile.png") %>" target="_blank">
			<img id="uploadPreview2" src="<% response.write("./images/nofile.png") %>" height="30px" width="25px"/><br/>
			</a>
<%
		End If
%>
		</td>
		<td colspan="4" style="padding-top: 8px; padding-left: auto; text-align: center; font-size: 1.0em; vertical-align: top;">
		</td>

	</tr>
<%
End If
%>

<?php
	$intervaldays = ceil(abs(strtotime($tglsd)- strtotime($tglmulai)) /(60*60*24));
?>
</table>

<table id="mrotab" name="mrotab" style="text-align: left; margin-left: 50px; margin-right: auto; margin-top: auto;" border="0" cellspacing="0" cellpadding="0">
	<tr height="20px" style="background-color: #FFFFFF">
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;" width="800px">
		</td>
	</tr>
	<tr height="20px" style="background-color: #FFFFFF">
		<td style="padding-left: auto; text-align: left; font-size: 1.0em;">
			<FONT face="Arial" color="#000">
<%
			Dim days as Long = 0
			days = DateDiff("d", Convert.ToDateTime(tglawalx), Convert.ToDateTime(tglakhirx))
%>
				Lama Waktu : &nbsp;&nbsp;<% response.write(days) %> &nbsp;hari kalendar
			</FONT>
		</td>
	</tr>
	<tr height="50px" style="background-color: #FFFFFF">
		<td style="padding-left: auto; text-align: center; font-size: 1.0em;">
		</td>
	</tr>
</table>


<form method="post" name="editkpfrm" id="editkpfrm" action="editpengadaannegosiasigo.php" enctype="multipart/form-data">
<table id="kpedittab" style="text-align: left; margin-left: 50px; margin-right: auto; margin-top: auto;" border="0" cellspacing="0" cellpadding="0">
	<!--tr height="20px">
		<td colspan="3" style="padding-left: 0px; text-align: left; font-size: 1.5em">
			<FONT face="Arial" color="#000"><b></b></FONT>
			<img src="./images/edit2.jpg" alt="" style="width:80px;height:80px;">
		</td>
	</tr-->
	<tr height="20px">
		<td colspan="3" style="padding-left: 20px; text-align: left;">
			<input type="hidden" id="edituserid" name="edituserid" value="<?php echo $userid;?>">
			<input type="hidden" id="editadaid" name="editadaid" value="<?php echo $adaid;?>">
		</td>
	</tr>
</table>	
</form>

<%
		End If
    Catch ex As Exception
        'lbl1.Text="Error while connecting Server."
		Response.Write (ex.Message)

    Finally
		'myReader.Close()
        'myConn.Close() 'Whether there is error or not. Close the connection.

    End Try

%>


<form name="viewijinfrm" id="viewijinfrm" action="permohonanedit.aspx" target="_blank" method="post" enctype="multipart/form-data">
	<input type="hidden" id="ijinuserid" name="ijinuserid" value="<% response.write(userid) %>">
	<input type="hidden" id="ijinadaid" name="ijinadaid" value="<% response.write(adaid) %>">
	<input type="hidden" id="ijinnext" name="ijinnext" value="1">
</form>

<form name="viewusulfrm" id="viewusulfrm" action="prosesusulan.aspx" target="_blank" method="post" enctype="multipart/form-data">
	<input type="hidden" id="usuluserid" name="usuluserid" value="<% response.write(userid) %>">
	<input type="hidden" id="usuladaid" name="usuladaid" value="<% response.write(adaid) %>">
	<input type="hidden" id="usulnext" name="usulnext" value="1">
</form>

<form name="viewaanwijfrm" id="viewaanwijfrm" action="prosesaanwijzing.aspx" target="_blank" method="post" enctype="multipart/form-data">
	<input type="hidden" id="aanuserid" name="aanuserid" value="<% response.write(userid) %>">
	<input type="hidden" id="aanadaid" name="aanadaid" value="<% response.write(adaid) %>">
	<input type="hidden" id="aannext" name="aannext" value="1">
</form>

<form name="viewpocfrm" id="viewpocfrm" action="prosespoc.aspx" target="_blank" method="post" enctype="multipart/form-data">
	<input type="hidden" id="pocuserid" name="pocuserid" value="<% response.write(userid) %>">
	<input type="hidden" id="pocadaid" name="pocadaid" value="<% response.write(adaid) %>">
	<input type="hidden" id="pocnext" name="pocnext" value="1">
</form>

<form name="viewevalfrm" id="viewevalfrm" action="prosesevaluasi.aspx" target="_blank" method="post" enctype="multipart/form-data">
	<input type="hidden" id="evaluserid" name="evaluserid" value="<% response.write(userid) %>">
	<input type="hidden" id="evaladaid" name="evaladaid" value="<% response.write(adaid) %>">
	<input type="hidden" id="evalnext" name="evalnext" value="1">
</form>

<form name="viewnegofrm" id="viewnegofrm" action="prosesnegosiasi.aspx" target="_blank" method="post" enctype="multipart/form-data">
	<input type="hidden" id="negouserid" name="negouserid" value="<% response.write(userid) %>">
	<input type="hidden" id="negoadaid" name="negoadaid" value="<% response.write(adaid) %>">
	<input type="hidden" id="negonext" name="negonext" value="1">
</form>

<form name="viewnkpfrm" id="viewnkpfrm" action="prosesnkp.aspx" target="_blank" method="post" enctype="multipart/form-data">
	<input type="hidden" id="nkpuserid" name="nkpuserid" value="<% response.write(userid) %>">
	<input type="hidden" id="nkpadaid" name="nkpadaid" value="<% response.write(adaid) %>">
	<input type="hidden" id="nkpnext" name="nkpnext" value="1">
</form>

<form name="viewsanggahfrm" id="viewsanggahfrm" action="prosessanggah.aspx" target="_blank" method="post" enctype="multipart/form-data">
	<input type="hidden" id="sanggahuserid" name="sanggahuserid" value="<% response.write(userid) %>">
	<input type="hidden" id="sanggahadaid" name="sanggahadaid" value="<% response.write(adaid) %>">
	<input type="hidden" id="sanggahnext" name="sanggahnext" value="1">
</form>

<form name="viewtunjukfrm" id="viewtunjukfrm" action="prosespenunjukan.aspx" target="_blank" method="post" enctype="multipart/form-data">
	<input type="hidden" id="tunjukuserid" name="tunjukuserid" value="<% response.write(userid) %>">
	<input type="hidden" id="tunjukadaid" name="tunjukadaid" value="<% response.write(adaid) %>">
	<input type="hidden" id="tunjuknext" name="tunjuknext" value="1">
</form>

<form name="viewspkfrm" id="viewspkfrm" action="prosesspk.aspx" target="_blank" method="post" enctype="multipart/form-data">
	<input type="hidden" id="spkuserid" name="spkuserid" value="<% response.write(userid) %>">
	<input type="hidden" id="spkadaid" name="spkadaid" value="<% response.write(adaid) %>">
	<input type="hidden" id="spknext" name="spknext" value="1">
</form>

<form name="viewpksfrm" id="viewpksfrm" action="prosespks.aspx" target="_blank" method="post" enctype="multipart/form-data">
	<input type="hidden" id="pksuserid" name="pksuserid" value="<% response.write(userid) %>">
	<input type="hidden" id="pksadaid" name="pksadaid" value="<% response.write(adaid) %>">
	<input type="hidden" id="pksnext" name="pksnext" value="1">
</form>


    <!-- using online scripts -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.6/umd/popper.min.js"
        integrity="sha384-wHAiFfRlMFy6i5SRaxvfOCifBUQy1xHdJ/yoi7FRNXMRBu5WHdZYu1hA6ZOblgut" crossorigin="anonymous">
    </script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.2.1/js/bootstrap.min.js"
        integrity="sha384-B0UglyR+jN6CkvvICOB2joaf5I4l3gm9GU6Hc1og6Ls7i6U/mkkaduKaBhlAXv9k" crossorigin="anonymous">
    </script>
    <script src="//malihu.github.io/custom-scrollbar/jquery.mCustomScrollbar.concat.min.js"></script>


<script>

    var config = {
      '.chosen-select'           : {},
      '.chosen-select-deselect'  : {allow_single_deselect:true},
      '.chosen-select-no-single' : {disable_search_threshold:10},
      '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
      '.chosen-select-width'     : {width:"65%"}
    }
    for (var selector in config) {
      $(selector).chosen(config[selector]);
    }

</script>


    <script src="pro-sidebar/src/js/main.js"></script>
	
<script type="text/javascript" src="datepicker/examples/public/javascript/jquery-1.11.1.js"></script>
<script type="text/javascript" src="datepicker/public/javascript/zebra_datepicker.js"></script>
<script type="text/javascript" src="datepicker/examples/public/javascript/core.js"></script>

<script>
	var jQuery190 = jQuery.noConflict();
	window.jQuery = jQuery190;
</script>

<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>
<script src="chosen1/select2.min.js"></script>
<script src="select2/dist/js/select2.min.js"></script>

<script>
$("#barupemohon").select2( {
 placeholder: "Pangkat/Golongan",
 allowClear: true
 } );
</script>
<script>
$("#barunip").select2( {
 placeholder: "NIP-Nama",
 allowClear: true
 } );
</script>
<script>
$("#editppk").select2( {
 placeholder: "PPK",
 allowClear: true
 } );
</script>

<script>
$("#editpemohon").select2( {
 placeholder: "Pangkat/Golongan",
 allowClear: true
 } );
</script>
<script>
$("#editnip").select2( {
 placeholder: "NIP-Nama",
 allowClear: true
 } );
</script>
<script>
$("#editkeputusan").select2( {
 placeholder: "Keputusan",
 allowClear: true
 } );
</script>


<script>
var $eventSelect = $("#barunamanip1");

$eventSelect.select2();

$eventSelect.on("change", function(e) {
    // insert your code here
	var nip = document.getElementById("barunamanip1").value;
//alert(nip);
	var userid = document.getElementsByName("kpuserid")[0].value;
	
});
</script>
	
</body>

</html>