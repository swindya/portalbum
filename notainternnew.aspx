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
	
<link rel="stylesheet" type="text/css" href="./tigracalendar/tcal.css" />
<script type="text/javascript" src="./tigracalendar/tcal.js"></script>

<!--link rel="stylesheet" href="css1/styles.css"-->

<link href="./datepicker4/dcalendar.picker.css" rel="stylesheet" type="text/css">

<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">



<style>
.example-element {
	margin-top: 0;
	position: fixed;
	background-color: #F9FBFB;
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
  float: left;
  display: block;
  color: #f2f2f2;
  text-align: center;
  padding: 17px 0px 17px 0px;
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

.welcome {
    left: 20px;
    line-height: 100px;
    margin-top: -40px;
	margin-left: 0px;
    position: absolute;
    text-align: left;
	color: #555555;
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
	color: #6681FF;
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
	font-size: 0.8em;
	color: #6D8FEB;
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
font-family: Calibri;

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
   margin: 2px;
   overflow: hidden;
   padding: 2px 8px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 90px;
}

select#soflow-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 10px;
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
   margin: 2px;
   overflow: hidden;
   padding: 2px 8px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 150px;
}

select#soflow1-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 10px;
}
select#soflow1A, select#soflow-color {
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
   margin: 2px;
   overflow: hidden;
   padding: 2px 8px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 150px;
}

select#soflow1A-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 10px;
}
select#soflow11, select#soflow11-color {
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
   padding: 2px 8px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 150px;
}

select#soflow11-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 10px;
}
select#soflow12, select#soflow12-color {
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
   padding: 2px 8px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 150px;
}

select#soflow12-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 10px;
}
select#soflow21, select#soflow21-color {
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
   padding: 2px 8px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 150px;
}

select#soflow21-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 10px;
}
select#soflow22, select#soflow22-color {
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
   padding: 2px 8px;
   text-overflow: ellipsis;
   white-space: nowrap;
   width: 150px;
}

select#soflow22-color {
   color: #fff;
   background-image: url(15xvbd5.png), -webkit-linear-gradient(#779126, #779126 40%, #779126);
   background-color: #779126;
   -webkit-border-radius: 20px;
   -moz-border-radius: 20px;
   border-radius: 20px;
   padding-left: 10px;
}
</style>

<style>
.css-button-3 {
	font-size: 12px;
	border-radius: 5px;
	border: solid 0px #18a418;
	color: #ffffff;
	background: linear-gradient(180deg, #32cd32 5%, #18a418 100%);
	box-shadow: 0px 5px 4px -2px #616174;
	//font-family: Arial;
	cursor: pointer;
	text-align: center;
	user-select: none;
	display: inline-flex;
	justify-content: center;
	align-items: center;
}
.css-button-3:hover {
	background: linear-gradient(180deg, #18a418 5%, #32cd32 100%);
}
.css-button-3:active {
	position: relative;
	top: 1px;
}
.css-button-3 > span {
	display: block;
}
.css-button-3-icon {
	padding: 6px 8px;
	border-right: 1px solid rgba(255, 255, 255, 0.16);
	box-shadow: rgba(0, 0, 0, 0.14) -1px 0px 0px inset;
}
.css-button-3-text {
	padding: 6px 8px;
}
</style>

<style>
.css-button-4 {
	font-size: 12px;
	border-radius: 5px;
	border: solid 0px #18a418;
	color: #ffffff;
	background: linear-gradient(180deg, #6085FF 5%, #0548a4 100%);
	box-shadow: 0px 5px 4px -2px #616174;
	//font-family: Arial;
	cursor: pointer;
	text-align: center;
	user-select: none;
	display: inline-flex;
	justify-content: center;
	align-items: center;
}
.css-button-4:hover {
	background: linear-gradient(180deg, #0548a4 5%, #6085FF 100%);
}
.css-button-4:active {
	position: relative;
	top: 1px;
}
.css-button-4 > span {
	display: block;
}
.css-button-4-icon {
	padding: 6px 8px;
	border-right: 1px solid rgba(255, 255, 255, 0.16);
	box-shadow: rgba(0, 0, 0, 0.14) -1px 0px 0px inset;
}
.css-button-4-text {
	padding: 6px 8px;
}
</style>

<style>
.container1 {
  margin: 20px 0px auto;
}
/* Ignore above this line */ 
 
input[type="checkbox"] {
  display: none;
}

input[type="checkbox"] + label span,input[type="checkbox"]:checked + label span {
  display: inline-block;
  width: 24px;
  height: 24px;
  margin: -20px 8px 0 0;
  vertical-align: top;
  cursor: pointer;
}

input[type="checkbox"] + label span {
  background: url('./images/variable-unchecked73.png') no-repeat;
}

input[type="checkbox"]:checked + label span {
  background: url('./images/variable-checked73.png') no-repeat;
} 
</style>

<style>

.container2 {
    margin: 0px 0px;
    width: 80px;
    text-align: left;
}
.container2 > .switch {
    display: block;
    margin: 0px 0px;
}
.switch {
    position: relative;
    display: inline-block;
    vertical-align: top;
    width: 62px;
    height: 20px;
    padding: 3px;
    background-color: white;
    border-radius: 18px;
    box-shadow: inset 0 -1px white, inset 0 1px 1px rgba(0, 0, 0, 0.05);
    cursor: pointer;
    background-image: -webkit-linear-gradient(top, #eeeeee, white 25px);
    background-image: -moz-linear-gradient(top, #eeeeee, white 25px);
    background-image: -o-linear-gradient(top, #eeeeee, white 25px);
    background-image: linear-gradient(to bottom, #eeeeee, white 25px);
}
.switch-input {
    position: absolute;
    top: 0;
    left: 0;
    opacity: 0;
}
.switch-label {
    position: relative;
    display: block;
    height: inherit;
    font-size: 10px;
    text-transform: uppercase;
    background: #eceeef;
    border-radius: inherit;
    box-shadow: inset 0 1px 2px rgba(0, 0, 0, 0.12), inset 0 0 2px rgba(0, 0, 0, 0.15);
    -webkit-transition: 0.15s ease-out;
    -moz-transition: 0.15s ease-out;
    -o-transition: 0.15s ease-out;
    transition: 0.15s ease-out;
    -webkit-transition-property: opacity background;
    -moz-transition-property: opacity background;
    -o-transition-property: opacity background;
    transition-property: opacity background;
}
.switch-label:before,
.switch-label:after {
    position: absolute;
    top: 50%;
    margin-top: -.5em;
    line-height: 1;
    -webkit-transition: inherit;
    -moz-transition: inherit;
    -o-transition: inherit;
    transition: inherit;
}
.switch-label:before {
    content: attr(data-off);
    right: 11px;
    color: #aaa;
    text-shadow: 0 1px rgba(255, 255, 255, 0.5);
}
.switch-label:after {
    content: attr(data-on);
    left: 11px;
    color: white;
    text-shadow: 0 1px rgba(0, 0, 0, 0.2);
    opacity: 0;
}
.switch-input:checked ~ .switch-label {
    background: #47a8d8;
    box-shadow: inset 0 1px 2px rgba(0, 0, 0, 0.15), inset 0 0 3px rgba(0, 0, 0, 0.2);
}
.switch-input:checked ~ .switch-label:before {
    opacity: 0;
}
.switch-input:checked ~ .switch-label:after {
    opacity: 1;
}
.switch-handle {
    position: absolute;
    top: 4px;
    left: 4px;
    width: 18px;
    height: 18px;
    background: white;
    border-radius: 10px;
    box-shadow: 1px 1px 5px rgba(0, 0, 0, 0.2);
    background-image: -webkit-linear-gradient(top, white 40%, #f0f0f0);
    background-image: -moz-linear-gradient(top, white 40%, #f0f0f0);
    background-image: -o-linear-gradient(top, white 40%, #f0f0f0);
    background-image: linear-gradient(to bottom, white 40%, #f0f0f0);
    -webkit-transition: left 0.15s ease-out;
    -moz-transition: left 0.15s ease-out;
    -o-transition: left 0.15s ease-out;
    transition: left 0.15s ease-out;
}
.switch-handle:before {
    content: '';
    position: absolute;
    top: 50%;
    left: 50%;
    margin: -6px 0 0 -6px;
    width: 12px;
    height: 12px;
    background: #f9f9f9;
    border-radius: 6px;
    box-shadow: inset 0 1px rgba(0, 0, 0, 0.02);
    background-image: -webkit-linear-gradient(top, #eeeeee, white);
    background-image: -moz-linear-gradient(top, #eeeeee, white);
    background-image: -o-linear-gradient(top, #eeeeee, white);
    background-image: linear-gradient(to bottom, #eeeeee, white);
}
.switch-input:checked ~ .switch-handle {
    left: 40px;
    box-shadow: -1px 1px 5px rgba(0, 0, 0, 0.2);
}
.switch-green > .switch-input:checked ~ .switch-label {
    background: #4fb845;
}

</style>


<script type="text/javascript">     
function PreviewImage(no) 
{       
	var namaim = "uploadImage" + no;
	var uploadim = "uploadPreview" + no;  
	var oFReader = new FileReader();
	var namafil = document.getElementById(namaim).value;
	var namafilearr = namafil.split(".");
	var noarr = namafilearr.length;
	var jenisfile = namafilearr[noarr-1];

//alert(namafil);	
	oFReader.readAsDataURL(document.getElementById(namaim).files[0]);
	
	if (jenisfile==='pdf')
	{
		oFReader.onload = function (oFREvent) {             
		document.getElementById(uploadim).src = './images/pdficon1.png'; 	
		};
	}
	else
	{
		oFReader.onload = function (oFREvent) {             
		document.getElementById(uploadim).src = oFREvent.target.result; 		
		};
	}
	//document.forms["memobarufrm"].submit();
	//memousulanbaru();

} 
</script>

<script>
function xxxhapussurat(id)
{
	if (confirm("Apakah anda ingin hapus data ?") == true) 
	{
		var tenure = prompt("Silahkan masukkan kode hapus", "");
		
		if (tenure == "opr1234") {
			var outx;
			var linkx = "hapuskctgo.aspx?id="+id;
			xhttp = new XMLHttpRequest();
			xhttp.onreadystatechange = function() {
				if (this.readyState == 4 && this.status == 200) {
					//document.getElementById("txtHint").innerHTML = this.responseText;
					outx = this.responseText;
					if (outx==1 || outx=="1")
					{
						alert("Data sudah dihapus.");
						window.location.reload();
					}
				}
			};
			xhttp.open("GET", linkx, true);
			xhttp.send();
		}
	}
}
</script>

<script>
function hapusnotin(id)
{
	if (confirm("Apakah anda ingin hapus data ?") == true) 
	{
		var tenure = prompt("Silahkan masukkan kode hapus", "");
		
		if (tenure == "123") {
			document.getElementsByName("delnotinid")[0].value = id;
			document.forms["delfrm"].submit();
		}
	}
}
</script>

<script>
function clearparam()
{
	document.getElementsByName("perihal")[0].value = "";
	document.getElementsByName("tglperiode1")[0].value = "";
	document.getElementsByName("tglperiode2")[0].value = "";
	document.forms["filterfrm"].submit();
}
</script>

<script>
function viewpage2(nohal)
{
var myval = document.getElementsByName("namanpp")[0].value;
if (myval==null || myval=="")
{
	var halink = "surat.aspx?page=" + nohal + "&cari=none";
}
else
{
	var halink = "surat.aspx?page=" + nohal + "&cari=" + myval;
}
//window.open(halink);
window.location.href = halink;
	//document.getElementsByName("pageno")[0].value = nohal;
	//document.forms["pagefrm"].submit();
}
</script>

<script>
function viewpage(nohal, totpage)
{
	if (nohal > totpage || nohal < 1)
	{
		return false;
	}
	document.getElementsByName("pageno")[0].value = nohal;
	document.getElementsByName("statusfilter")[0].value = 1;
	document.forms["filterfrm"].submit();
}
</script>

<script>
function viewfilter()
{
//var myval = document.getElementsByName("namanpp")[0].value;
//var halink = "main.aspx?page=1&cari=" + myval;
//window.open(halink);
//window.location.href = halink;
	document.getElementsByName("statusfilter")[0].value = 1;
	document.forms["filterfrm"].submit();
}
</script>

<script>
function databaru()
{
	document.getElementById("databarutab").style.display = "block";
	document.getElementById("dataedittab").style.display = "none";
	document.getElementById("filtertab").style.display = "none";
	document.getElementById("barubtntab").style.display = "none";
	document.getElementById("datatab").style.display = "none";
	document.getElementById("totalpagetab").style.display = "none";
	document.getElementById("pagingtab").style.display = "none";
	//document.getElementById("clearbtn").style.display = "none";
	//document.getElementById("filtertab").style.display = "none";
}
</script>

<script>
function hidedatabaru()
{
	document.getElementById("databarutab").style.display = "none";
	document.getElementById("dataedittab").style.display = "none";
	document.getElementById("filtertab").style.display = "block";
	document.getElementById("barubtntab").style.display = "block";
	document.getElementById("datatab").style.display = "block";
	document.getElementById("totalpagetab").style.display = "block";
	document.getElementById("pagingtab").style.display = "block";
	//document.getElementById("clearbtn").style.display = "none";
	//document.getElementById("filtertab").style.display = "none";
}
</script>

<script>
function hidedataedit()
{
	document.getElementById("databarutab").style.display = "none";
	document.getElementById("dataedittab").style.display = "none";
	document.getElementById("filtertab").style.display = "block";
	document.getElementById("barubtntab").style.display = "block";
	document.getElementById("datatab").style.display = "block";
	document.getElementById("totalpagetab").style.display = "block";
	document.getElementById("pagingtab").style.display = "block";
	//document.getElementById("clearbtn").style.display = "none";
	//document.getElementById("filtertab").style.display = "none";
}
</script>

<script>
function hideeditbaru()
{
	document.getElementById("databarutab").style.display = "none";
	document.getElementById("dataedittab").style.display = "none";
	document.getElementById("filtertab").style.display = "block";
	document.getElementById("barubtntab").style.display = "block";
	document.getElementById("datatab").style.display = "block";
	document.getElementById("totalpagetab").style.display = "block";
	document.getElementById("pagingtab").style.display = "block";
	//document.getElementById("clearbtn").style.display = "none";
	//document.getElementById("filtertab").style.display = "none";
}
</script>

<script>
function savebaru()
{
	var tgl = document.getElementsByName("barutgl")[0].value;
	var jenis = document.getElementsByName("barujenis")[0].value;
	var perihal = document.getElementsByName("baruperihal")[0].value;
	var tujuan = document.getElementsByName("barutujuan")[0].value;
	var pengirim = document.getElementsByName("barupengirim")[0].value;
	if ((tgl==null || tgl=="") || (jenis==null || jenis=="") || (perihal==null || perihal=="") || (tujuan==null || tujuan=="") || (pengirim==null || pengirim==""))
	{
		alert("Mohon lengkapi data terlebih dahulu !");
		return false;
	}

	document.getElementsByName("baruju")[0].value = 0;
	document.getElementsByName("baruada")[0].value = 0;
	var und = document.getElementById("barujunda");
	var ada = document.getElementById("barupengadaan");
	if (und.checked==true)
	{
		document.getElementsByName("baruju")[0].value = 1;
	}
	if (ada.checked==true)
	{
		document.getElementsByName("baruada")[0].value = 1;
	}
	
	document.forms["barufrm"].submit();
}
</script>

<script>
function saveedit()
{
	var abc = document.getElementById("uploadImage1").value;
//alert(abc);
	document.getElementsByName("editju")[0].value = 0;
	document.getElementsByName("editada")[0].value = 0;
	var und = document.getElementById("editjunda");
	var ada = document.getElementById("editpengadaan");
	if (und.checked==true)
	{
		document.getElementsByName("editju")[0].value = 1;
	}
	if (ada.checked==true)
	{
		document.getElementsByName("editada")[0].value = 1;
	}
	document.forms["editfrm"].submit();
}
</script>

<script>
function editdatax(id, nonota, tgl, jenis, perihal, rubrik, tujuan, tembusan, jenisundangan, statuspengadaan, namafile, namafolder, extens, statusbum, wrnama, wrnpp, namaklp)
{

	document.getElementById("databarutab").style.display = "none";
	document.getElementById("dataedittab").style.display = "block";
	document.getElementById("filtertab").style.display = "none";
	document.getElementById("barubtntab").style.display = "none";
	document.getElementById("datatab").style.display = "none";
	document.getElementById("totalpagetab").style.display = "none";
	document.getElementById("pagingtab").style.display = "none";

	document.getElementsByName("editnotinid")[0].value = id;
	//document.getElementsByName("editnomor")[0].value = nonota;
	document.getElementsByName("edittgl")[0].value = tgl;
	document.getElementsByName("editperihal")[0].value = perihal;
	document.getElementsByName("edittujuan")[0].value = tujuan;
	document.getElementsByName("edittembusan")[0].value = tembusan;
	
	document.getElementById('nomor').innerHTML = nonota;
	document.getElementById('namapp').innerHTML = "<i>Petugas Input :</i> &nbsp;&nbsp;" + wrnama + " (" + wrnpp + ") - " + namaklp;

	//document.getElementById("soflow21").value = jenis;
	//document.getElementById("soflow22").value = rubrik;
	document.getElementsByName("oldnamafile")[0].value = namafile;
	if (jenisundangan==1)
	{
		document.getElementById("editjunda").checked = true;
	}
	else
	{
		document.getElementById("editjunda").checked = false;
	}
	
	if (statuspengadaan==1)
	{
		document.getElementById("editpengadaan").checked = true;
	}
	else
	{
		document.getElementById("editpengadaan").checked = false;
	}	
	
	var folderfile = "." + namafolder + namafile;
//alert(extens);
	//document.getElementById("uploadImage1").value = namafile;
	document.getElementById("uploadPreview1").src = folderfile;
	//document.getElementsByName("uploadImage0")[0].value = gambar;
	document.getElementById("linkimage1").href = folderfile;
	if (extens=="pdf")
	{
		document.getElementById("uploadPreview1").src = "./images/pdficon1.png";
	}
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
	Dim a, aaa, b, bb, c, d, p, q, jmla, jmldata as Long
	Dim user0 as string
	Dim pwd0 as string
	Dim namauserx as String
	Dim aksesx as String
	Dim idz(200000) as String
	Dim tglzz(50000) as String
	Dim idx(50000) as String
	Dim nosuratx(50000) as String
	Dim tglx(50000) as String
	Dim jenisx(50000) as String
	Dim perihalx(50000) as String
	Dim tujuanx(50000) as String
	Dim tembusanx(50000) as String
	Dim jenisundanganx(50000) as String
	Dim cuserx(50000) as String
	Dim namafilex(50000) as String
	Dim namafolderx(50000) as String
	Dim rubrikx(1000) as String
	Dim rubriks as String
	Dim jmlrubkel as Integer = 0
	
	Dim idq As String = ""
	Dim namauserq As String = ""
	Dim statusq As String = ""
	Dim nppq as String = ""
	Dim jabatanq As String = ""
	Dim klpq As String = ""
	Dim unitq As String = ""
	Dim posisiq As String = ""
	Dim suratq As String = ""
	Dim sekreq As String = ""
	Dim adaq as String
	Dim asetq as String
	Dim statusaktifq as String = ""
	Dim divq As String = ""
	Dim levelaksesq as String = ""
	Dim posisix as String = ""
	Dim jmlkel as Integer = 0
	
	Dim idxx as String
	Dim notinxx as String
	Dim tglxx as String
	Dim jenisxx as String
	Dim rubrikxx as String
	Dim perihalxx as String
	Dim tujuanxx as String
	Dim tembusanxx as String
	Dim jenisundanganxx as String
	Dim statusadaxx as String
	Dim cuserxx as String
	Dim namafilexx as String
	Dim namafolderxx as String
	Dim namaklpq as String = ""
	Dim nourutxx as String = ""
	Dim jmlrubriko as String = ""
	
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
	Dim ii as Long = 0
	Dim k as Long = 0
	Dim userid as Integer = 1
	Dim statusbum as Integer = 0
	Dim namauserstr as String = ""
	Dim namaklpxx as String = ""
	Dim writenama as String = ""
	Dim writenpp as String = ""
	Dim klpxx as Integer = 0

	Dim SQL0 As String
	Dim SQL1 As String
	Dim SQL2 As String
	Dim SQLUP as String
	
	Dim startval as Integer = 0
	Dim endval as Integer = 0
	Dim datastrx(5000) as String
	Dim datastrxx as String
	Dim perihal as String = ""
	Dim perihalstr as String = ""
	Dim tgl1 as String = ""
	Dim tgl2 as String = ""
	Dim tgl10 as string = ""
	Dim tgl20 as String = ""
	Dim tglstr as String = ""
	Dim tglsu as String
	Dim statusf as Integer = 0
	Dim i as Long
	Dim l as Integer
	Dim m as integer
	Dim n as Integer
	Dim jmlrubrik as Integer = 0
	Dim klpstr as String = ""
	Dim rubarr(10) as String
	Dim statusbaru as Integer = 0
	Dim namakel(100) as String
	Dim jmlrubiko as Integer = 0
	Dim rubriko(100) as String
	Dim rubkel(3000) as String

	
	Dim setmenu as String = "notin"
	
	Dim tahuniki as Integer
	Dim tahunwingi as Integer
	Dim tahunsesuk as Integer

	'lbl1.Text = Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
	tahuniki = Val(Year(Now))
	tahunwingi = tahuniki - 1
	tahunsesuk = tahuniki + 1
	
	aaa=1
	
	hal = 1
	pageno = 1
	'GET
	'If NOT Request.QueryString("pageno") is Nothing Then
	'	 pageno = Request.QueryString("pageno")
	'	 hal = Cint(pageno)
	'End If
	

    Try
        If aaa=1 Then
			'myConn = New SqlConnection("Data Source=DESKTOP-S7N8DF9;Initial Catalog=KCT;User ID=sa;Password=sa;")
			pagerows = (hal-1) * 20
%>

<!--#include file="koneksi.aspx"-->

<%			
	
			If Request.Form("pageno") isNot Nothing Then
				pageno = Request.Form("pageno")
				hal = Cint(pageno)
			End If
			If NOT Request.Form("perihal") is Nothing Then
				perihal = Request.Form("perihal")
			End If
			If NOT Request.Form("tglperiode1") is Nothing Then
				tgl1 = Request.Form("tglperiode1")
			End If
			If NOT Request.Form("tglperiode2") is Nothing Then
				tgl2 = Request.Form("tglperiode2")
			End If
			If NOT Request.Form("statusfilter") is Nothing Then
				statusf = Cint(Request.Form("statusfilter"))
			End If
    
			user0 = Session("username")
			idq = Session("userid")
			nppq = Session("idemployee")
			If not Session("namauser") is nothing Then
				namauserq = Session("namauser")
			End If
			jabatanq = Session("jabatan")
			klpq = Session("klp")
			If not Session("namaklp") is nothing Then
				namaklpq = Session("namaklp")
			End If
			unitq = Session("unit")
			posisiq = Session("posisi")
			posisix = Session("posisistr")
			suratq = Session("menusurat")
			sekreq = Session("menusekre")
			adaq = Session("menuada")
			asetq = Session("menuaset")
			divq = Session("divisi")
			levelaksesq = Session("levelakses")

			namauserstr = namauserq
			If String.IsNullOrEmpty(namauserq) or namauserq.length()<2 Then
			'If namauserq.length()<2 Then
%>

<script>
alert('Anda tidak beraktivitas terlalu lama (idle) atau sesi anda habis. Silahkan login lagi');
window.top.location.href = "logout.aspx"; 
</script>

<%
			Else If namauserq.length()> 14 Then
				namauserstr = namauserq.Substring(0, 14)	
			End If
			myCmd = myConn.CreateCommand

'If perihal IsNot Nothing || NOT perihal Is Nothing
			perihalstr = ""
			If perihal.length()>0 Then
				perihalstr = " AND (perihal LIKE '%" & perihal & "%' OR tujuan LIKE '%" & perihal & "%')"
			End If
			
			tglstr = ""
			Dim tglarr() as String
			If tgl1.length() > 1 AND String.IsNullOrEmpty(tgl2) Then
				tglarr = tgl1.Split("-")
				tgl10 = tglarr(2) & "-" & tglarr(1) & "-" & tglarr(0)
				tglstr = " AND tglnotin='" & tgl10 & "'"
			Elseif String.IsNullOrEmpty(tgl1) AND tgl2.length() > 1 Then
				tglarr = tgl2.Split("-")
				tgl20 = tglarr(2) & "-" & tglarr(1) & "-" & tglarr(0)
				tglstr = " AND tglnotin='" & tgl20 & "'"
			Elseif tgl1.length()>1 AND tgl2.length()>1 Then
				tglarr = tgl1.Split("-")
				tgl10 = tglarr(2) & "-" & tglarr(1) & "-" & tglarr(0)
				tglarr = tgl2.Split("-")
				tgl20 = tglarr(2) & "-" & tglarr(1) & "-" & tglarr(0)
				tglstr = " AND tglnotin BETWEEN '" & tgl10 & "' AND '" & tgl20 & "'"
			End If

			klpstr = " AND klp=" & klpq
'============================================================================
' Data notin
'============================================================================
			SQL0 = "SELECT ID FROM [notin] WHERE (ID>0" & perihalstr & tglstr & klpstr & ") ORDER BY tglnotin DESC"
			If namaklpq.ToUpper() = "BUM" Then
				SQL0 = "SELECT ID FROM [notin] WHERE (ID>0" & perihalstr & tglstr & ") ORDER BY tglnotin DESC"
			End If

			myCmd.CommandText = SQL0
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			
			a = 0
			jmldata = 0
            If myReader.HasRows Then
                While myReader.Read()
					a = a + 1
					idz(a) = myReader("ID").ToString()
					'tglzz(a) = myReader("tglnotin1").ToString()
				End While
			End If
			myConn.Close()
			jmldata = a

'==========================================================================
' Cari Nama Divisi
'==========================================================================
			Dim namadiv as String = ""
			SQL0 = "SELECT * FROM [divisi] WHERE (div_id=" & divq & ")"

			myCmd.CommandText = SQL0
			myConn.Open()
			myReader = myCmd.ExecuteReader()
            If myReader.HasRows Then
                While myReader.Read()
					namadiv = myReader("div_nama").ToString()
				End While
			End If
			myConn.Close()
'==========================================================================
' Cari Nama Kelompok: SEMUA
'==========================================================================
			SQL0 = "SELECT * FROM [klp] ORDER BY klp_nama ASC"
			d = 0
			myCmd.CommandText = SQL0
			myConn.Open()
			myReader = myCmd.ExecuteReader()
            If myReader.HasRows Then
                While myReader.Read()
					d = d + 1
					namakel(d) = myReader("klp_nama").ToString()
				End While
			End If
			myConn.Close()
			jmlkel = d
'==========================================================================
' Cari Daftar Rubrik dari divisinya BUM
'==========================================================================
			'SQL0 = "SELECT * FROM [rubrik] WHERE (divisi=" & divq & " AND kelompok=" & klpq & ") ORDER BY rubrik ASC"
			SQL0 = "SELECT * FROM [rubrik] WHERE (divisi=" & divq & ") ORDER BY rubrik ASC"
			If namaklpq.ToUpper() = "BUM" Then
				SQL0 = "SELECT * FROM [rubrik] WHERE (divisi=" & divq & ") ORDER BY rubrik ASC"
				statusbum = 1
			End If
			myCmd.CommandText = SQL0
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			b = 0
            If myReader.HasRows Then
                While myReader.Read()
					b = b + 1
					rubrikx(b) = myReader("rubrik").ToString()
					rubriks = rubrikx(b)
					If rubriks.length() > 1 Then
						rubarr = rubriks.Split("/")
						If rubarr.count() > 2 Then
							rubrikx(b) = namaklpq & "/" & rubarr(1) & "/"
							rubriko(b) = "/" & rubarr(1) & "/"
						Else
							b = b - 1
						End If
					End If
				End While
			End If

			myConn.Close()

			jmlrubrik = b
			jmlrubriko = b

'==========================================================================
			p = 0
			For c = 1 to jmlkel
				For d = 1 to jmlrubriko
					p = p + 1
					rubkel(p) = namakel(c) & rubriko(d)
				Next d
			Next c
			jmlrubkel = p
'==========================================================================
' Cari Daftar Rubrik dari divisinya user
'==========================================================================
			'SQL0 = "SELECT * FROM [rubrik] WHERE (divisi=" & divq & " AND kelompok=" & klpq & ") ORDER BY rubrik ASC"
			SQL0 = "SELECT * FROM [rubrik] WHERE (divisi=" & divq & " AND kelompok=" & klpq & ") ORDER BY rubrik ASC"
			If namaklpq.ToUpper() = "BUM" Then
				SQL0 = "SELECT * FROM [rubrik] WHERE (divisi=" & divq & ") ORDER BY rubrik ASC"
				statusbum = 1
			End If
			myCmd.CommandText = SQL0
			myConn.Open()
			myReader = myCmd.ExecuteReader()
			b = 0
            If myReader.HasRows Then
                While myReader.Read()
					b = b + 1
					rubrikx(b) = myReader("rubrik").ToString()
					rubriks = rubrikx(b)
					If rubriks.length() > 1 Then
						rubarr = rubriks.Split("/")
						If rubarr.count() > 2 Then
							rubrikx(b) = namaklpq & "/" & rubarr(1) & "/"
						Else
							b = b - 1
						End If
					End If
				End While
			End If

			myConn.Close()

			jmlrubrik = b

'==========================================================================

'** Cari jumlah halaman
			if (jmldata >= rowperpage) Then
				b = jmldata / rowsperpage
				c = jmldata / rowsperpage
				jmlpage = c
				if (b > c) then
					jmlpage = c + 1
				End If
			else
				jmlpage = 1
			End If

			If hal = jmlpage Then
				rowsperpage = jmldata - ((jmlpage - 1)*20)
			End If

			startval = ((hal-1)*20) + 1
			endval = startval + rowsperpage - 1
'*********

        End If

    Catch ex As Exception
        'lbl1.Text="Error while connecting Server."
		Response.Write (ex.Message)

    Finally
		'myReader.Close()
        'myConn.Close() 'Whether there is error or not. Close the connection.

    End Try
%>
<div class="example-element"></div>

<!--#include file="menu.aspx"-->

<!--#include file="topbar.aspx"-->


<div class="container">
			<div class="welcome">Selamat Datang,</div>
			<div class="photo"><img src="./images/user5.png" alt="" style="width:40px;height:40px;"></div>
			<div class="userin"><i><b><% Response.Write(namauserstr) %></b></i></div>
			<div class="posisi" STYLE="font-size: 0.8em"><i><% Response.Write(posisix) %></i></div>
</div>

<table style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: auto;" width="950px" border="0" cellspacing="0" cellpadding="0">
	<tr height="30px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		</td>
	</tr>
	<tr height="20px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 1.5em">
			<FONT face="Arial" color="#666666">DAFTAR NOTA INTERN</FONT>
		</td>
	</tr>
	<tr height="50px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		
		</td>
	</tr>
</table>

<form method="post" name="barufrm" id="barufrm" action="notainternaddgo.aspx" enctype="multipart/form-data">
<table id="databarutab" style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: 0; display: none" width="1050px" border="0" cellspacing="0" cellpadding="0">
	<tr height="0px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		</td>
	</tr>
	<tr height="20px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 1.4em">
			<FONT face="Arial" color="#44AA44">NOTA INTERN BARU</FONT>
		</td>
	</tr>
	<tr height="40px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top" width="220px">
			<FONT face="Arial" color="#000">
				Tgl 
			</FONT>
			<FONT face="Arial" color="#DD0000">
				* 
			</FONT>				
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top" width="30px">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top" width="700px">
			<FONT face="Arial" color="#000">
				<input style="padding-left:10px; font-size: 0.9em; background-color: #FFF; border:1px solid #ababab; height: 20px; width: 100px" autocomplete="off" type="text" name="barutgl" id="barutgl" class="tcal"  STYLE="color: #000000; background-color: #FFFFFF;" tabindex="1"/>
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em" width="100px">										
		</td>
	</tr>
	<!--tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Tahun Pengadaan
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.0em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 130px" id="soflow2" name="barutahun" data-placeholder="Jenis" tabindex="2"> 
				<option value="" >&nbsp;&nbsp;</option>
				<option value="<% Response.Write(tahunwingi)%>"><% Response.Write(tahunwingi)%></option>
				<option value="<% Response.Write(tahuniki)%>"><% Response.Write(tahuniki)%></option>
				<option value="<% Response.Write(tahunsesuk)%>"><% Response.Write(tahunsesuk)%></option>
			</select>			
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr-->
	<!--tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Jenis
			</FONT>
			<FONT face="Arial" color="#DD0000">
				* 
			</FONT>				
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.0em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 130px" id="soflow11" name="barujenis" data-placeholder="Jenis" tabindex="2"> 
				<option value="" >&nbsp;&nbsp;</option>
				<option value="1" >Biasa</option>
				<option value="2" >Rahasia</option>
			</select>			
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr-->
	<tr height="70px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Perihal
			</FONT>
			<FONT face="Arial" color="#DD0000">
				* 
			</FONT>				
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.0em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<textarea STYLE="font-family: Arial; height: 55px; width: 640px; color: #000000; padding-left: 5px; padding-middle: 5px; padding-bottom: 5px; border:1px solid #ababab; font-weight: normal; font-size:1.0em; line-height: 110%;" id="baruperihal" name="baruperihal" tabindex="3"></textarea>
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Pengirim (Rubrik)
			</FONT>
			<FONT face="Arial" color="#DD0000">
				* 
			</FONT>	
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
<%
		If namaklpq.ToUpper() = "BUM" Then
%>
			<div id="page-wrapper">
			  
				<input STYLE="color: #000000; font-family: Arial; font-weight: normal; font-size:0.9em; text-align: left; background-color: #ffffff; width: 120px; height:20px;" class="textboxwide" type="text" id="barupengirim" name="barupengirim" list="languages" placeholder="Rubrik">
				<datalist id="languages">
<%
			For p = 1 to jmlrubkel
%>
    <option value="<% response.write(rubkel(p)) %>"><% response.write(rubkel(p)) %></option>
<%
			Next p
%>
				</datalist>
			</div>
<%
		Else
%>

			<FONT face="Arial" color="#000">
				<!--input style="font-family: Arial; font-size: 1.0em; width: 560px" type="text" id="barupengirim" name="barupengirim" class="textbox" onchange="getnama(this.value);" tabindex="4"-->
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 120px; height: 25px" id="soflow12" name="barupengirim" data-placeholder="Jenis" tabindex="4"> 
				<option value="" >&nbsp;&nbsp;</option>
<%
			for b=1 to jmlrubrik
%>
				<option value="<% response.write(rubrikx(b)) %>" ><% response.write(rubrikx(b)) %></option>
<%
			next b
%>
			</select>		
			</FONT>
<%
		End If
%>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Tujuan
			</FONT>
			<FONT face="Arial" color="#DD0000">
				* 
			</FONT>	
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<input style="font-size: 0.9em; width: 640px; height: 18px" type="text" id="barutujuan" name="barutujuan" class="textbox" onchange="getnama(this.value);" tabindex="5">
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Tembusan
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<input style="font-size: 0.9em; width: 640px; height: 18px" type="text" id="barutembusan" name="barutembusan" class="textbox" onchange="getnama(this.value);" tabindex="6">
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Undangan
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<div style="text-align:left;margin-top:0%">
					<div class="container2">
						<label class="switch switch-green">
						<input type="checkbox" class="switch-input" id="barujunda">
						<span class="switch-label" data-on="Yes" data-off="No"></span>
						<span class="switch-handle"></span>
						</label>
					</div>
				</div>
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Pengadaan
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<div style="text-align:left;margin-top:0%">
					<div class="container2">
						<label class="switch switch-green">
						<input type="checkbox" class="switch-input" id="barupengadaan" name="barupengadaan">
						<span class="switch-label" data-on="Yes" data-off="No"></span>
						<span class="switch-handle"></span>
						</label>
					</div>
				</div>
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; padding-top: 5px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Upload Dokumen
			</FONT>											
		</td>
		<td style="padding-left: 0px; padding-top: 5px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: auto; padding-top: 5px; text-align: left; font-size: 1.2em; vertical-align: top">
			<table id="file1" style="text-align: left; margin-left: 0px; margin-right: auto; margin-top: auto;" border="0" cellspacing="0" cellpadding="0">
				<tr height="42px">
					<td style="padding-left: 0px; padding-top: 0px; vertical-align: top; text-align: left; font-size: 1.0em">
						<FONT face="Arial" color="#000">
							<input id="uploadImage0" name="uploadImage0" type="file" onchange="PreviewImage(0);" onclick="resetpic();" tabindex="6"  />
						</FONT>
					</td>
					<td style="padding-top: 5px; padding-left: auto; text-align: right; font-size: 1.2em" width="40px">
						<FONT face="Arial" color="#000">
							<a id="linkimage0" href="#" target="_blank">
							<img id="uploadPreview0" height="100px" width="80px"/><br/>
							</a>
						</FONT>
					</td>
				</tr>
			</table>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="20px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 0.8em">
			<FONT face="Arial" color="#DD0000">
				 * Harus diisi (mandatory)
			</FONT>		
		</td>
	</tr>
	<tr height="40px">
		<td colspan="4" style="padding-left: 20px; text-align: left;">
		</td>
	</tr>
	<tr height="40px">
		<td style="padding-left: 0px; text-align: left;">
			<a id="batalbtn" href="" class="medium button orange" onclick="hidedatabaru();">Tutup</a>
		</td>
		<td colspan="3" style="padding-left: 0px; text-align: left;">
			<a id="simpanbtn" href="#" class="medium button blue" onclick="savebaru();"><i class="fa fa-save" aria-hidden="true"></i>&nbsp;&nbsp;Simpan</a>
		</td>
	</tr>
	<tr height="30px">
		<td colspan="4">
			<input type="hidden" id="baruju" name="baruju">
			<input type="hidden" id="baruada" name="baruada">
			<input type="hidden" id="barujenis" name="barujenis" value="1">
			<input type="hidden" id="editfile" name="editfile">
		</td>
	</tr>
</table>
</form>

<%
'response.write("-----------------------------------" & namauserx)
%>

<form method="post" name="editfrm" id="editfrm" action="notainterneditgo.aspx" enctype="multipart/form-data">
<table id="dataedittab" style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: 0; display: none" width="1050px" border="0" cellspacing="0" cellpadding="0">
	<tr height="0px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		</td>
	</tr>
	<tr height="20px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 1.3em">
			<FONT face="Arial" color="#045CFF">VIEW/EDIT DATA</FONT>
		</td>
	</tr>
	<tr height="40px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top" width="220px">
			<FONT face="Arial" color="#000">
				Tgl 
			</FONT>
			<FONT face="Arial" color="#DD0000">
				* 
			</FONT>		
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top" width="30px">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top" width="700px">
			<FONT face="Arial" color="#000">
				<input style="padding-left:10px; font-size: 0.9em; background-color: #FFF; border:1px solid #ababab; height: 26px; width: 140px" autocomplete="off" type="text" name="edittgl" id="edittgl" class="tcal"  STYLE="color: #000000; background-color: #FFFFFF;" tabindex="1"/>
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em" width="100px">										
		</td>
	</tr>
	<!--tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Jenis
			</FONT>
			<FONT face="Arial" color="#DD0000">
				* 
			</FONT>		
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.0em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 130px" id="soflow21" name="editjenis" data-placeholder="Jenis" tabindex="2"> 
				<option value="" >&nbsp;&nbsp;</option>
				<option value="1" >Biasa</option>
				<option value="2" >Rahasia</option>
			</select>			
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr-->
	<tr height="70px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Perihal
			</FONT>
			<FONT face="Arial" color="#DD0000">
				* 
			</FONT>		
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.0em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<textarea STYLE="font-family: Arial; height: 55px; width: 640px; color: #000000; padding-left: 5px; padding-middle: 5px; padding-bottom: 5px; border:1px solid #ababab; font-weight: normal; font-size:0.9em; line-height: 110%;" id="editperihal" name="editperihal" rows="2" cols="70" tabindex="3"></textarea>
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<!--tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Pengirim (Rubrik)
			</FONT>
			<FONT face="Arial" color="#DD0000">
				* 
			</FONT>					
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
			<select STYLE="color: #000000; background-color: #ffffff; font-size: 0.9em; width: 130px" id="soflow22" name="editpengirim" data-placeholder="Pengirim" tabindex="4"> 
				<option value="" >&nbsp;&nbsp;</option>
<%
			for b=1 to jmlrubrik
%>
				<option value="<% response.write(rubrikx(b)) %>" ><% response.write(rubrikx(b)) %></option>
<%
			next b
%>
			</select>		
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr-->
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Nomor
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#4444FF">
				<label id="nomor" style="font-size: 1.0em; font-weight: 600"></label>			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Tujuan
			</FONT>
			<FONT face="Arial" color="#DD0000">
				* 
			</FONT>					
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<input style="font-size: 0.9em; width: 640px; height: 18px" type="text" id="edittujuan" name="edittujuan" class="textbox" onchange="getnama(this.value);" tabindex="5">
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Tembusan
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<input style="font-size: 0.9em; width: 640px; height: 18px" type="text" id="edittembusan" name="edittembusan" class="textbox" onchange="getnama(this.value);" tabindex="6">
			</FONT>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Undangan
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<div style="text-align:left;margin-top:0%">
					<div class="container2">
						<label class="switch switch-green">
						<input type="checkbox" class="switch-input" id="editjunda">
						<span class="switch-label" data-on="Yes" data-off="No"></span>
						<span class="switch-handle"></span>
						</label>
					</div>
				</div>
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Pengadaan
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Arial" color="#000">
				<div style="text-align:left;margin-top:0%">
					<div class="container2">
						<label class="switch switch-green">
						<input type="checkbox" class="switch-input" id="editpengadaan" name="editpengadaan">
						<span class="switch-label" data-on="Yes" data-off="No"></span>
						<span class="switch-handle"></span>
						</label>
					</div>
				</div>
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; padding-top: 5px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				Attach Dok
			</FONT>											
		</td>
		<td style="padding-left: 0px; padding-top: 5px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Arial" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: auto; padding-top: 5px; text-align: left; font-size: 1.2em; vertical-align: top">
			<table id="file1" style="text-align: left; margin-left: 0px; margin-right: auto; margin-top: auto;" border="0" cellspacing="0" cellpadding="0">
				<tr height="42px">
					<td style="padding-left: 0px; padding-top: 0px; vertical-align: top; text-align: left; font-size: 1.0em">
						<FONT face="Arial" color="#000">
							<input id="uploadImage1" name="uploadImage1" type="file" onchange="PreviewImage(1);" onclick="resetpic();" tabindex="6"  />
						</FONT>
					</td>
					<td style="padding-top: 5px; padding-left: auto; text-align: right; font-size: 1.2em" width="40px">
						<FONT face="Arial" color="#000">
							<a id="linkimage1" href="#" target="_blank">
							<img id="uploadPreview1" height="100px" width="80px"/><br/>
							</a>
						</FONT>
					</td>
				</tr>
			</table>
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 1.2em">										
		</td>
	</tr>
	<tr height="20px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 0.8em">
			<FONT face="Arial" color="#DD0000">
			</FONT>		
		</td>
	</tr>
	<tr height="20px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 0.9em">
			<FONT face="Arial" color="#000000">
				<label id="namapp"></label>
			</FONT>		
		</td>
	</tr>
	<tr height="20px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 0.8em">
			<FONT face="Arial" color="#DD0000">
			</FONT>		
		</td>
	</tr>
	<tr height="20px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 0.8em">
			<FONT face="Arial" color="#DD0000">
				 * Harus diisi (mandatory)
			</FONT>		
		</td>
	</tr>
	<tr height="40px">
		<td colspan="4" style="padding-left: 20px; text-align: left;">
		</td>
	</tr>
	<tr height="40px">
		<td style="padding-left: 0px; text-align: left;">
			<a id="batalbtn" href="" class="medium button orange" onclick="hidedataedit();">Tutup</a>
		</td>
		<td colspan="3" style="padding-left: 0px; text-align: left;">
			<a id="simpanbtn" href="#" class="medium button blue" onclick="saveedit();"><i class="fa fa-save" aria-hidden="true"></i>&nbsp;&nbsp;Simpan</a>
		</td>
	</tr>
	<tr height="30px">
		<td colspan="4">
			<input type="hidden" id="editnotinid" name="editnotinid">
			<input type="hidden" id="oldnamafile" name="oldnamafile">
			<input type="hidden" id="editju" name="editju">
			<input type="hidden" id="editada" name="editada">
			<input type="hidden" id="editfile" name="editfile">
			<input type="hidden" id="editjenis" name="editjenis" value="1">
		</td>
	</tr>
</table>
</form>

<table style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: auto;" width="950px" border="0" cellspacing="0" cellpadding="0">
	<tr height="0px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 1.6em">
		</td>
	</tr>
</table>

<form method="post" name="filterfrm" id="filterfrm" action="notaintern.aspx" enctype="multipart/form-data">
<table id="filtertab" style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: auto;" width="1150px" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFEDD5">
	<tr height="20px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; padding-left: 30px; text-align: left; font-size: 0.9em; vertical-align: top" width="190px">
			<FONT face="Calibri" color="#000">
				Periode 
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top" width="40px">
			<FONT face="Calibri" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top" width="920px">
			<FONT face="Calibri" color="#000">
				<input style="padding-left:10px; font-size: 0.9em; background-color: #FFF; border:1px solid #ababab; height: 20px; width: 100px" autocomplete="off" type="text" name="tglperiode1" id="tglperiode1" class="tcal"  STYLE="color: #000000; background-color: #FFFFFF;" tabindex="1"/>
			&nbsp;&nbsp; s/d &nbsp;&nbsp;
				<input style="padding-left:10px; font-size: 0.9em; background-color: #FFF; border:1px solid #ababab; height: 20px; width: 100px" autocomplete="off" type="text" name="tglperiode2" id="tglperiode2" class="tcal"  STYLE="color: #000000; background-color: #FFFFFF;" tabindex="1"/>
			</FONT>
		</td>
	</tr>
	<tr height="35px">
		<td style="padding-left: 0px; padding-left: 30px; text-align: left; font-size: 0.9em; vertical-align: top">
			<FONT face="Calibri" color="#000">
				Perihal / Tujuan
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: center; font-size: 0.9em; vertical-align: top">
			<FONT face="Calibri" color="#000">
				 :
			</FONT>											
		</td>
		<td style="padding-left: 0px; text-align: left; vertical-align: top">
			<FONT face="Calibri" color="#000">
				<!--select id="barutpmobil" name="barutpmobil" STYLE="color: #000000; background-color: #ffffff; font-size: 1.2em; width:320px"-->
				<div>
					<!--input style="padding-right: 10px; text-align: left; font-family: Calibri; font-size: 1.2em" type="text" id="brgmerk" name="brgmerk" class="textbox"-->
					<input style="padding-right: 10px; text-align: left; font-size: 0.9em; width: 400px" type="text" id="perihal" name="perihal" class="textbox" value="<% Response.Write(perihal) %>">
				</div>
			</FONT>
		</td>
	</tr>
	<tr height="20px">
		<td colspan="3" style="padding-left: 20px; text-align: left;">
		</td>
	</tr>
	<tr height="40px">
		<td style="padding-left: 30px; text-align: left;">
			<a id="caribtn" href="#" class="medium button blue" onclick="viewfilter();"><i class="fas fa-search" aria-hidden="true"></i>&nbsp;&nbsp;Cari</a>
		</td>
		<td colspan="2" style="padding-left: 30px; text-align: left;">
			<a id="clearbtn" href="#" class="medium button red" onclick="clearparam();">Reset</a>
		</td>
	</tr>
	<tr height="20px">
		<td colspan="3" style="padding-left: 20px; text-align: left;">
			<input type="hidden" id="userid" name="userid" value="">
			<input type="hidden" id="pageno" name="pageno" value="1">
			<input type="hidden" id="statusfilter" name="statusfilter" value="0">
		</td>
	</tr>
</table>
</form>

<table style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: auto;" width="860px" border="0" cellspacing="0" cellpadding="0">
	<tr height="30px">
		<td colspan="4" style="padding-left: 20px; text-align: left; font-size: 1.0em">
		</td>
	</tr>
</table>

<table id="barubtntab" style="margin-left: 270px;" width="1150px" align="center" cellpadding="2" cellspacing="2" border="0" bgcolor="#EEE" >
	<tr height="38px" align="left">
		<td style="font-family: Arial; font-size: 0.9em;text-align: left;" width="120px">
			<!--a id="barubbm" href="" class="small button blue" onclick="addkct(); return false;"><i class="fas fa-plus fa-sm"></i>&nbsp;&nbsp;New</a-->
			<a href="#" class="css-button-3" style="text-decoration: none; font-size: 1.1em" onclick="databaru();">
				<span class="css-button-3-icon"><i class="fa fa-plus" aria-hidden="true" id="i_pos"></i></span>
				<span class="css-button-3-text"><FONT face="Calibri" size="2">Data Baru</FONT></span>
			</a>
		</td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: left;" width="690px">
		</td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: left;" width="140px">
<%
If namaklpq.ToUpper() = "BUM" Then
%>
			<a href="#" class="css-button-4" style="text-decoration: none; font-size: 1.1em" onclick="databaru();">
				<span class="css-button-4-icon"><i class="fa fa-plus" aria-hidden="true" id="i_pos"></i></span>
				<span class="css-button-4-text"><FONT face="Calibri" size="2">Export to CSV</FONT></span>
			</a>
<%
End If
%>
		</td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: left;" width="60px">
		</td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: left;" width="140px">
			<FONT face="Calibri" color="#000"><b>Jml data:&nbsp;<% response.write(jmldata.ToString("#,#")) %></b></FONT>
		</td> 		
	</tr>
</table>
<%
'response.write("-------------------------------------------" & SQL0 & "<br>")
'response.write("-------------------------------------------" & jmldata & "<br>")
			
'response.write("---------------------------------------------------------" & SQL1 & vbcrlf)
'response.write("---------------------------------------------------------" & pageno & "--" & hal & "   start:" & startval & "-" & endval)

%>
<table id="datatab" style="margin-left: 270px;" width="1150px" align="center" cellpadding="2" cellspacing="2" border="0" bgcolor="#FAFAFA" >
	<tr height="42px" align="left" style="background-color:#D5D5D5;color:Black;" >
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="60px"><b>NO</b></td>                        
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="110px"><b>TGL NOTA</b></td>            
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="120px"><b>NO NOTA</b></td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="400px"><b>PERIHAL</b></td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="310px"><b>TUJUAN</b></td>
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="75px"><b>EDIT</b></td>
<%
If namaklpq.ToUpper() = "BUM" Then
%>
		<td style="font-family: Arial; font-size: 0.9em;text-align: center;" width="75px"><b>HAPUS</b></td>		
<%
End If
%>
	</tr>

<%

	'm = 46 MOD 2
	'response.write("-------------------------------------------------" & m & " start:" & startval & "--" & endval)
Try
	ii = pagerows
	k = 1
	a = endval - startval
	bb = 0
	n = 0
	for i = startval to endval
		ii = ii + k
		bb = bb + 1
		SQL1 = "SELECT * FROM notin WHERE (ID=" & idz(i) & ")"
		'SQL1 = "SELECT * FROM notin WHERE ((ID>=10 AND ID<=50)" & _
		'		") ORDER BY tglnotin DESC"
				
		myCmd.CommandText = SQL1
		myConn.Open()
		myReader = myCmd.ExecuteReader()
		
		a = i
		If myReader.HasRows Then
			While myReader.Read()
				idxx = myReader("ID").ToString()
				notinxx = myReader("nomor").ToString()
				tglxx = myReader("tglnotin").ToString()
				jenisxx = myReader("jenis").ToString()
				rubrikxx = myReader("rubrik").ToString()
				nourutxx = myReader("nourut").ToString()
				klpxx = Val(myReader("klp").ToString())
				perihalxx = myReader("perihal").ToString()
				tujuanxx = myReader("tujuan").ToString()
				tembusanxx = myReader("cc").ToString()
				jenisundanganxx = myReader("undangan").ToString()
				statusadaxx = myReader("statuspengadaan").ToString()
				cuserxx = myReader("createduser").ToString()
				namafilexx = myReader("namafile").ToString()
				namafolderxx = myReader("namafolder").ToString()
			End While
		End If
		myConn.Close()
'================================================================================
		SQL1 = "SELECT * FROM [user] WHERE ID=" & cuserxx			
		myCmd.CommandText = SQL1
		myConn.Open()
		myReader = myCmd.ExecuteReader()
		

		If myReader.HasRows Then
			While myReader.Read()
				writenama = myReader("nama").ToString()
				writenpp = myReader("idemployee").ToString()
			End While
		End If
		myConn.Close()
		
		SQL1 = "SELECT * FROM [klp] WHERE klp_id=" & klpxx			
		myCmd.CommandText = SQL1
		myConn.Open()
		myReader = myCmd.ExecuteReader()
		

		If myReader.HasRows Then
			While myReader.Read()
				namaklpxx = myReader("klp_nama").ToString()
			End While
		End If
		myConn.Close()
'================================================================================		
		
		notinxx = rubrikxx & nourutxx
		
		SQLUP = "UPDATE notin SET nomor='" & rubrikxx & nourutxx & "' WHERE ID=" & idz(i)
		myCmd.CommandText = SQLUP
		myConn.Open()
		myReader = myCmd.ExecuteReader()
		myConn.Close()
		



		'Dim oDate As DateTime = DateTime.Parse(tglx(i))
		'tglsu = oDate.Day & " " & oDate.Month & "  " & oDate.Year
		'MsgBox(oDate.Day & " " & oDate.Month & "  " & oDate.Year)

		Dim dt As DateTime = Convert.ToDateTime(tglxx)
		'tglsu = dt.Day & " " & dt.Month & "  " & dt.Year
		tglsu = format(dt, "d-M-yyyy")
		'Dim dt As DateTime = format(tglefektifx(i), "yyyy-MM-dd")
		'tglefektifx(i) = dt
		'tglarr= tglx(i).Split("-")
		'tglx(i) = tglarr(2) & "-" & tglarr(1) & "-" & tglarr(0)
		
		'datastrx(i) = idx(i) & ",'" & nonotinx(i) & "','" & tglx(i) & "','" & jenisx(i) & "','" & perihalx(i) & "','" & tujuanx(i) & "','" & _
		'				tembusanx(i) & "','" & jenisundanganx & "','" & namafilex(i) & "','" & namafolderx(a) & "'"
		Dim extens as String = ""		
		If namafilexx.length() > 1 Then
			Dim filearr() as String
			filearr = namafilexx.Split(".") 
			extens = filearr(1)
		End If
				
		
		datastrxx = idxx & ",'" & notinxx & "','" & tglsu & "','" & jenisxx & "','" & perihalxx & "','" & rubrikxx & "','" & tujuanxx & "','" & _
						tembusanxx & "','" & jenisundanganxx & "','" & statusadaxx & "','" & namafilexx & "','" & namafolderxx & "','" & _
						extens & "'," & statusbum & ",'" & writenama & "','" & writenpp & "','" & namaklpxx & "'"
						
		'datastrx(i) = idx(i) & "," & namax(i)
%>
<%
	If n=0 Then
		n = 1
%>
	<tr id="row<% Response.Write(bb)%>" STYLE="background-color: #FFFFFF" height="38px">
<%
	Elseif n=1 Then
		n = 0
%>
	<tr id="row<% Response.Write(bb)%>" STYLE="background-color: #FFF1E3" height="38px">
<%
	End If
%>
		<td style="padding-right: 5px; padding-top: 5px;  font-family: Arial; font-size: 0.9em;text-align: right; vertical-align: top; line-height: 110%"><% Response.Write (i) %></td>
		<td style="padding-left: 5px; padding-top: 5px;  font-family: Arial; font-size: 0.9em;text-align: center; vertical-align: top; line-height: 110%"><% Response.Write (tglsu) %></td>
		<td style="padding-left: 5px; padding-top: 5px;  font-family: Arial; font-size: 0.9em;text-align: left; vertical-align: top; line-height: 110%"><% Response.Write (notinxx) %></td>
		<td style="padding-left: 5px; padding-top: 5px;  font-family: Arial; font-size: 0.9em;text-align: left; vertical-align: top; line-height: 110%"><% Response.Write (perihalxx) %></td>
		<td style="padding-left: 5px; padding-top: 5px;  font-family: Arial; font-size: 0.9em;text-align: left; vertical-align: top; line-height: 110%"><% Response.Write (tujuanxx) %></td>
		<!--td style="padding-left: 0px; font-family: Arial; font-size: 0.9em;text-align: center;">
			<a id="linkimage<% response.write(i) %>" href="<% Response.Write(namafilexx)%>" target="_blank">
			<img id="uploadPreview<% response.write(i) %>" src="<% Response.Write(namafilexx)%>" height="26px" width="25px"/><br/>
			</a>
		</td-->
		<td style="padding-left: 0px; padding-top: 5px; font-family: Arial; font-size: 0.9em; text-align: center; vertical-align: top;">
			<!--a id="edit<% Response.Write (i)%>" href="" class="small button blue" onclick="addkct(); return false;"><i class="fas fa-edit fa-sm"></i></a-->
			<a id="edit<% Response.Write (i)%>" href="#" onclick="editdatax(<% Response.Write(datastrxx) %>);">
			<img id="uploadPreview<% Response.Write (i)%>" src="./images/edit2.jpg" height="26px" width="25px"/><br/>
			</a>
		</td>
<%
If namaklpq.ToUpper() = "BUM" Then
%>
		<td style="padding-left: 0px; padding-top: 5px; font-family: Arial; font-size: 0.9em;text-align: center; vertical-align: top;">
			<!--a id="hapus<% Response.Write(i)%>" href="#" class="small button red" onclick="addkct(); return false;"><i class="fas fa-trash fa-sm"></i></a-->		
			<a id="hapus<% Response.Write (i)%>" href="#" onclick="hapusnotin(<% Response.Write(idxx) %>);">
			<img id="uploadPreview<% Response.Write (i)%>" src="./images/delete4.png" height="26px" width="25px"/><br/>
			</a>
		</td>
<%
End If
%>
	</tr>
<%
	next

Catch ex As Exception
        'lbl1.Text="Error while connecting Server."
		Response.Write (ex.Message)

Finally
		'myReader.Close()
        'myConn.Close() 'Whether there is error or not. Close the connection.	
End Try
%>

</table>

<table style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: auto;" width="1140px" border="0" cellspacing="0" cellpadding="0">
	<tr height="30px">
		<td colspan="4" style="padding-left: 0px; text-align: left; font-size: 1.4em">

		</td>
	</tr>
</table>

<table id="totalpagetab" style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: auto;" width="1140px" border="0" cellspacing="0" cellpadding="0">
	<tr height="30px">
		<td style="padding-left: 0px; text-align: left; font-size: 0.9em">
			<FONT face="Calibri" color="#000"><b>Total Hal:&nbsp;<% response.write(jmlpage) %></b></FONT>
		</td>
	</tr>
	<tr height="10px">
		<td style="padding-left: 0px; text-align: left; font-size: 1.4em">
		</td>
	</tr>
</table>

<table id="pagingtab" style="text-align: left; margin-left: 270px; margin-right: auto; margin-top: auto;" width="1100px" border="0" cellspacing="0" cellpadding="0">
	<tr height="34px" style="background-color: #FFFFFF">
		<td style="padding-left: auto; text-align: left; font-size: 1.1em;">
			<div class="pagination">
				<a href="#" onclick="viewpage(<% Response.Write(1) %>, <% Response.Write(jmlpage) %>)">&lt;&lt;</a>
				<a href="#" onclick="viewpage(<% Response.Write(pageno-1) %>, <% Response.Write(jmlpage) %>)">&lt;</a>
<%
			m = pageno \ 20
			n = pageno MOD 20
			startpage = (m * 20) + 1
			if n=0 AND m>0 Then
				startpage = ((m-1) * 20) + 1
			End If
			endpage = startpage + 19

			If endpage > jmlpage Then
				endpage = jmlpage
			End If
			
			for a=startpage to endpage
%>
			  <a 
			  <% If a=pageno Then 
					Response.Write ("class='active'")
				 End If
			  %>
				href="#" onclick="viewpage(<% Response.Write(a) %>, <% Response.Write(jmlpage) %>)"><% Response.Write(a) %></a>
<%
			next a
%>
				<a href="#" onclick="viewpage(<% Response.Write(pageno+1) %>, <% Response.Write(jmlpage) %>)">&gt;</a>
				<a href="#" onclick="viewpage(<% Response.Write(jmlpage) %>, <% Response.Write(jmlpage) %>)">&gt;&gt;</a>
			</div>
			
		</td>
		<td style="padding-left: 0px; text-align: left; font-size: 1.0em; vertical-align: middle;">
			
		</td>
	</tr>
	<tr height="30px">
		<td colspan="2" style="padding-left: 0px; text-align: left; font-size: 1.4em">
		</td>
	</tr>
</table>

<form method="post" name="delfrm" id="delfrm" action="notainterndelgo.aspx" enctype="multipart/form-data">
	<input type="hidden" id="delnotinid" name="delnotinid">
</form>

<%
			Session("username") = user0
			Session("userid") = idq
			Session("namauser") = namauserq
			Session("idemployee") = nppq
			Session("levelakses") = levelaksesq
			Session("klp") = klpq
			Session("namaklp") = namaklpq
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
</body>
</html>
