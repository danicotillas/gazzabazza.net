function calendar(formname,currentdate)
	{
	self.name = 'opener';
	remote = open('calendar.asp?name=' + formname + '&sdate=' + currentdate, 'remote', 'width=160,height=165,location=no,scrollbars=no,menubars=no,toolbars=no,resizable=yes,fullscreen=no');
 	remote.focus();
	}

function MM_preloadImages() 
	{
	  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
	    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
	    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
	}
	
function MM_swapImgRestore() 
	{
	  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
	}
	
function MM_findObj(n, d) 
	{
	  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
	    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
	  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
	  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
	  if(!x && d.getElementById) x=d.getElementById(n); return x;
	}
	
function MM_swapImage() 
	{
	  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
	   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
	}
	
function showhide(id)
	{
		var itm = null;
		if (document.getElementById) {
			itm = document.getElementById(id);
		} else if (document.all){
			itm = document.all[id];
		} else if (document.layers){
			itm = document.layers[id];
		}
		if (itm.style) {
			if (itm.style.display == "none") { itm.style.display = ""; }
			else { itm.style.display = "none"; }
		}
	}

function showhideconfig(id,show)
		{
			var itm = null;
			if (document.getElementById) {
				itm = document.getElementById(id);
			} else if (document.all){
				itm = document.all[id];
			} else if (document.layers){
				itm = document.layers[id];
			}
			if (itm.style) {
				if (itm.style.display == "none" && show == 1) 
					{ 
					itm.style.display = ""; 
					}
				if (itm.style.display == "" && show == 0) 
					{ 
					itm.style.display = "none"; 
					}
			}
		}

function printpreview()
	{
	showhide('header');
	showhide('chooser');
	showhide('pgfooter');
	showhideconfig('about',0);
	}

function SetCookie (name,value,expires,path,domain,secure) 
	{
	document.cookie = name + "=" + escape (value) +
	((expires) ? "; expires=" + expires.toGMTString() : "") +
	((path) ? "; path=" + path : "") +
	((domain) ? "; domain=" + domain : "") +
	((secure) ? "; secure" : "");
	}

	
function hideabout()
	{
	showhide('about');
	SetCookie ("about", "HIDE", null);
	}
	
function exportreport(mode)
	{
		document.exportform.action = "export.asp?type=" + mode;
		document.exportform.submit();
		return false;
	}
	
function submitwhoisquery(registry, ipaddress)
	{
	if (registry == "ARIN")
		{
		document.arin.queryinput.value = ipaddress;
		document.arin.submit();
		}
	if (registry == "APNIC")
		{
		document.apnic.searchtext.value = ipaddress;
		document.apnic.submit();
		}
	if (registry == "RIPE")
		{
		document.ripe.searchtext.value = ipaddress;
		document.ripe.submit();
		}
	if (registry == "LACNIC")
		{
		document.lacnic.query.value = ipaddress;
		document.lacnic.submit();
		}
	if (registry == "AFRINIC")
		{
		document.afrinic.searchtext.value = ipaddress;
		document.afrinic.submit();
		}
	}

function validateform(form,formid)
		{
			if (formid == 3)
				{
				for (var i = 0; i < form.action.length; i++)
					{
					if (form.action[i].checked)
						{
						var straction = form.action[i].value
						break
			      		}
			  		}
				
				if (straction == "delete")
					{
					var agree=agreedelete();
					if (agree)
						return true;
					else
						return false;
					}
					
				if (straction == "definitions" || straction == "countries")
					{ 
					if (form.file.value == ''){
						var agree=agreefile(form, straction);
						if (agree)
							return true;
						else
							return false;
						}
					}
			}
			return true;
		}
	
function agreedelete()
		{
		var agree=confirm("Deleting all statistics data is permanent. Are you sure you wish to proceed?");
		if (agree)
			return true;
		else
			return false;
		}
		
function agreesetup(intType)
		{
		if (intType == 1)
			var agree=confirm("The FJstats install will take a few minutes to load the nessasary data. Press OK to continue.");
		else
			var agree=confirm("Upgrading FJstats from a previous version can take a long time depending on the amount of data selected to upgrade. Press OK to continue or CANCEL to change your setup options.");
		if (agree)
			return true;
		else
			return false;
		}

function agreefile(form, action)
		{
		var agree=confirm("You have not selected a file to upload. By proceeding, FJstats will try and use the data\\" 
		+ action + ".txt file in the FJstats installation folder.");
		if (agree)
			return true;
		else
			return false;
		}
		
function checkdates(form)
		{
		if (form.start.value == ''){
			alert('Please enter a start date.');
			return false;
			form.start.focus();
			}
		if (form.end.value == ''){
			alert('Please enter an end date.');
			return false;
			form.end.focus();
			}
		}
		
function setupform()
		{
		var install = document.setup.install;
		var dbtype = document.setup.dbtype;
		var db2type = document.setup.db2type;
		var dbprefix = document.setup.dbprefix;
		var dbcreate = document.setup.dbcreate;
		
		if (install.options[install.selectedIndex].value == 2)
			{
			showhideconfig('upgrade',1);
			}
		else
			{
			showhideconfig('upgrade',0);
			}
		if (dbtype.options[dbtype.selectedIndex].value == "MSACCESS")
			{
			dbprefix.value = 'mt_';
			dbcreate.checked = false;
			showhideconfig('dbusername',0);
			showhideconfig('dbpassword',0);
			showhideconfig('dbprefix',0);
			showhideconfig('dbcreate',0);
			}
		else
			{
			dbcreate.checked = true;
			showhideconfig('dbusername',1);
			showhideconfig('dbpassword',1);
			showhideconfig('dbprefix',1);
			showhideconfig('dbcreate',1);
			}	
		if (db2type.options[db2type.selectedIndex].value == "MSACCESS")
			{
			document.setup.db2prefix.value = '';
			showhideconfig('db2username',0);
			showhideconfig('db2password',0);
			showhideconfig('db2prefix',0);
			}
		else
			{
			showhideconfig('db2username',1);
			showhideconfig('db2password',1);
			showhideconfig('db2prefix',1);
			}	
		}
function showhelp(helpfile,bookmark)
		{
		self.name = 'opener';
		helpwin = open('hlp' + helpfile + '.htm#' + bookmark, 'helpwin', 'width=400,height=500,location=no,scrollbars=yes,menubars=no,toolbars=no,resizable=yes,fullscreen=no');
 		helpwin.focus();
		}