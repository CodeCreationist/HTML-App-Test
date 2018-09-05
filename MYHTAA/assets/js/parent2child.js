//this one doesnt use accessdb.js I needed more fucntionality so I did it myself
//this loads the index.hta information
function loadProcessesss()
{
	try{
		var adoConn = new ActiveXObject("ADODB.Connection");
		var adoRS = new ActiveXObject("ADODB.Recordset");
		adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
		var processesFA = document.getElementById("ProcessesFA");
		var fillarea = "";
		adoRS.Open("Select * From HomePage WHERE Hidden='F'", adoConn);
		if(!adoRS.bof)
		{
			adoRS.MoveFirst;
			while(!adoRS.eof)
			{
				fillarea += "<section id='"+adoRS.fields(0)+"' class='wrapper spotlight style1'>" ;
				fillarea += "<div class='inner'>";
				fillarea += "<a onclick='return setcookieredir("+adoRS.fields(0)+")' class='image'><img src='images/form.png' alt='' /></a>";
				fillarea += "<div class='content'>";
				fillarea += "<h2 class='major'>"+adoRS.fields(1)+"</h2><p>"+adoRS.fields(2)+"</p>";
				fillarea += "<a onclick='return setcookieredir("+adoRS.fields(0)+")'  class='special'>Learn more</a>";
				fillarea += "</div></div></section>";
				adoRS.movenext;
			}
			processesFA.innerHTML = fillarea;
			adoRS.Close();
			adoConn.Close();
		}
		isadmin();
		loadattachments();
	}
	catch(e)
	{
		var e1 = "parent2child.js had an error in loadProcessesss() function";
		ErrorToFile(e, e1);
		location.reload();
	}
};
function setCookie(cname, cvalue, exdays) {
	try{
		var d = new Date();
		d.setTime(d.getTime() + (exdays*24*60*60*1000));
		var expires = "expires="+d.toUTCString();
		document.cookie = cname + "=" + cvalue + "; " + expires;
	}catch(e){ErrorToFile(e);}
}

function getCookie(cname) {
	try{
		var name = cname + "=";
		var ca = document.cookie.split(';');
		for(var i = 0; i < ca.length; i++) {
			var c = ca[i];
			while (c.charAt(0) == ' ') {
				c = c.substring(1);
			}
			if (c.indexOf(name) == 0) {
				return c.substring(name.length, c.length);
			}
		}
		return "";
	}catch(e){ErrorToFile(e);}
}

function checkCookie() {
	try{
		var user = getCookie("username");
		if (user != "") {
			alert("Welcome again " + user);
		} else {
			user = prompt("Please enter your name:", "");
			if (user != "" && user != null) {
				setCookie("username", user, 365);
			}
		}
	}catch(e){ErrorToFile(e);}
}
function setcookieredir(id){
	try{
		setCookie("ID", id, 1);
		window.location = "processviewer.hta";
	}catch(e){ErrorToFile(e);}
}
function returnID(){
	return getCookie("ID");
}
//checks if the user is an admin 
function checkadmin(){
	try{
		if(isadmin() != true)
			window.location = "index.hta";
		if(deleteenabled() == false)
			document.getElementById("DAT").style.display = "none";
	}catch(e){ErrorToFile(e);}
}
function isadmin()
{
	try{
		var WinNetwork = new ActiveXObject("WScript.Network");
		var myDB = new ACCESSdb("Abcon.mdb", {showErrors:true});
		var SQL = "Select * From AdminTable Where AdminName='"+WinNetwork.UserName+"'";
		if(myDB.query(SQL))
		{
			return true;
		}
		else{
			if(document.getElementById("leftnave") != null)
				document.getElementById("leftnave").style.display = 'none';
			return false;
		}
	}
	catch(e)
	{
		var e1 = "parent2child.js had an error in isadmin() function";
		ErrorToFile(e, e1);
		location.reload();
	}
}
//admin option saves a cookie and checks 
function toggledelete(){
	try{
		if(getCookie("Del") != 1)
			setCookie("Del", 1, 1);
		else
			setCookie("Del", 0, 1);
	}catch(e){ErrorToFile(e);}
}
//checks if delete is enabled
function deleteenabled(){
	try{
		if(getCookie("Del") != 1)
			return false;
		else
			return true;
	}catch(e){ErrorToFile(e);}
}
function loadattachments()
{
	try{
		var adoConn = new ActiveXObject("ADODB.Connection");
		var adoRS = new ActiveXObject("ADODB.Recordset");
		adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
		var TableFA = document.getElementById("TablesFA");
		var fillarea2 = "";
		adoRS.Open("Select * From Attachments", adoConn);
		if(!adoRS.bof)
		{
			adoRS.MoveFirst;
			while(!adoRS.eof)
			{
				fillarea2 += "<article><input type='hidden' id='Grab#"+adoRS.fields(0)+"' value='"+adoRS.fields("Att").value+"'/>";
				fillarea2 += "<a onclick='dlatt("+adoRS.fields(0)+")' class='image'><img src='images/dlatt.png' alt='' /></a>";
				fillarea2 += "<h3 class='major'>"+adoRS.fields("FileName").value+"</h3><a onclick='dlatt("+adoRS.fields(0)+")' class='special'>Learn more</a></article>";
				adoRS.movenext;
			}
			TableFA.innerHTML = fillarea2;
			adoRS.Close();
			adoConn.Close();
		}
	}
	catch(e)
	{
		var e1 = "parent2child.js had an error in loadattachments() function";
		ErrorToFile(e, e1);
		location.reload();
	}
	
}
//doesnt use accessdb.js
function loadTables()
{
	try{
		setCookie("Search", 0, 1);
		var adoConn = new ActiveXObject("ADODB.Connection");
		var adoRS = new ActiveXObject("ADODB.Recordset");
		adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
		var processesFA = document.getElementById("ProcessesFA");
		var TableFA = document.getElementById("TablesFA");
		var fillarea = "";
		var fillarea2 = "";
		var ID = returnID();
		adoRS.Open("Select * From TableList WHERE ParentProcessID="+ID+" AND Obsolete='F'", adoConn);
		if(!adoRS.bof)
		{
			adoRS.MoveFirst;
			while(!adoRS.eof)
			{
				fillarea += "<section id='"+adoRS.fields(0)+"' class='wrapper spotlight style1'>" ;
				fillarea += "<div class='inner'>";
				fillarea += "<a href='#pop4form' onmouseover='return generatefill("+adoRS.fields(0)+")' id='ID#"+adoRS.fields(0)+"' name='"+adoRS.fields(1)+"' class='image'><img src='images/form.png' alt='' /></a>";
				fillarea += "<div class='content'>";
				fillarea += "<h2 class='major'>"+adoRS.fields(1)+"</h2>";
				fillarea += "<a href='#pop4form' onmouseover='return generatefill("+adoRS.fields(0)+")' class='special'>Learn more</a>";
				fillarea += "</div></div></section>";
				fillarea2 += "<article><a href='#pop4table' onmouseover='return generateTable("+adoRS.fields(0)+")' id='ID#"+adoRS.fields(0)+"' name='"+adoRS.fields(1)+"' class='image'><img src='images/Table.png' alt='' /></a>";
				fillarea2 += "<h3 class='major'>"+adoRS.fields(1)+"</h3><p>View Table "+adoRS.fields(1)+".</p><a href='#pop4table' onmouseover='return generateTable("+adoRS.fields(0)+")'class='special'>Learn more</a></article>";
				adoRS.movenext;
			}
			processesFA.innerHTML = fillarea;
			TableFA.innerHTML = fillarea2;
			adoRS.Close();
			adoConn.Close();
		}
	}
	catch(e)
	{
		var e1 = "parent2child.js had an error in loadTables() function";
		ErrorToFile(e, e1);
		location.reload();
	}
};
function makememo(ID, i)
{
	try{
		var Grab = document.getElementById("Grab#"+ID+"N#"+i+"");
		var getFill = document.getElementById("fillp4f");
		var temp = "<h2>"+Grab.name+"</h2><blockquote>"+Grab.value+"</blockquote><a id='close' href='#' class='close'>&nbsp;&nbsp;&nbsp;</a>"
		getFill.innerHTML = temp;
		$body = $('body');
		$body.addClass('is-pop4form-visible');
		return false;
	}catch(e){ErrorToFile(e);}
}
//fills up the create table javascript
function editrow(ID, TID, Del){
	try{
		generatefill(TID);
		var itstxt = document.getElementById("TxtFill");
		var itsMemo = document.getElementById("MemoFill");
		var itsFill = document.getElementById("Auto-Fill");
		var itsHyper = document.getElementById("hyperfill");
		var itsLong = document.getElementById("LongFill");
		var itsDou = document.getElementById("DouFill");
		var itsCurr = document.getElementById("CurrFill");
		var itsDate = document.getElementById("DateFill");
		var itsCheck = document.getElementById("CheckFill");
		var itsDrop = document.getElementById("DropFill");
		var adoConn = new ActiveXObject("ADODB.Connection");
		var adoRS = new ActiveXObject("ADODB.Recordset");
		var TableName = document.getElementById("ID#"+TID+"").name;
		adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
		adoRS.Open("Select * From "+TableName+" Where ID="+ID+"", adoConn, 1, 3);
		
		for(var i = 1; i<= itsCheck.name; i++)
		{
			var tempY = document.getElementById("CheckYFill"+i+"");
			var tempN = document.getElementById("CheckNFill"+i+"");
			tempN = false;
			if(adoRS.Fields(tempY.name).value == 1)
				tempY.checked = true;
			else
				tempN = true;
		}
		for(var i = 1; i<= itsHyper.name; i++)
		{
			var temp = document.getElementById("HyperFill"+i+"");
			temp.innerHTML = (adoRS.Fields(temp.alt).value);
		}
		for(var i = 1; i<= itsMemo.name; i++)
		{
			var temp = document.getElementById("MemoFill"+i+"");
			if(adoRS.Fields(temp.name).value != null)
				temp.value = adoRS.Fields(temp.name).value;
		}
		for(var i = 1; i<= itstxt.name; i++)
		{
			var temp = document.getElementById("TxtFill"+i+"");
			temp.value = adoRS.Fields(temp.name).value;
		}
		for(var i = 1; i<= itsFill.name; i++)
		{
			var temp = document.getElementById("AutoF#"+i+"");
			temp.value = adoRS.Fields(temp.name).value;
		}
		for(var i = 1; i<= itsLong.name; i++)
		{
			var temp = document.getElementById("LongFill"+i+"");
			temp.value = adoRS.Fields(temp.name).value;
		}
		for(var i = 1; i<= itsDou.name; i++)
		{
			var temp = document.getElementById("DouFill"+i+"");
			temp.value = adoRS.Fields(temp.name).value;
		}
		for(var i = 1; i<= itsCurr.name; i++)
		{
			var temp = document.getElementById("CurrFill"+i+"");
			temp.value = adoRS.Fields(temp.name).value;
		}
		for(var i = 1; i<= itsDate.name; i++)
		{
			var temp = document.getElementById("DateFill"+i+"");
			var tempval = (new Date((adoRS.Fields(temp.name).value))).format("MM/dd/yyyy");
			temp.value = tempval;
		}
		for(var i = 1; i<= itsDrop.name; i++)
		{
			var temp = document.getElementById("DropFill"+i+"");
			temp.options[0].text = adoRS.Fields(temp.name).value;
		}
		var changeoptions = document.getElementById("changeoptions");
		if(Del == null)
		{
			changeoptions.innerHTML = "<ul class='actions'>"+
				"<li class='six columns'><input type='submit' onclick='return createrecord4autofill("+TID+", "+ID+")' value='Edit Information' class='special' /></li>"+
				"<li class='six columns'><input type='reset' value='Reset Form' /></li>"+
			"</ul>";
		}
		else
		{
			changeoptions.innerHTML = "<ul class='actions'>"+
				"<li class='six columns'><input type='submit' onclick='return createrecord4autofill("+TID+", "+ID+", "+ID+")' value='Delete Information' class='special' /></li>"+
				"<li class='six columns'><input type='reset' value='Reset Form' /></li>"+
			"</ul>";
		}
		$body = $('body');
		$body.addClass('is-pop4form-visible');
		adoRS.Close();
		adoConn.Close();
	}
	catch(e)
	{
		var e1 = "parent2child.js had an error in editrow(ID, TID, Del) function";
		ErrorToFile(e, e1);
		location.reload();
	}
}
function addsearch2t(ID)
{
	try{
		var max4table = getCookie("HMS");
		var addsearch2t = document.getElementById("addsearches2t");
		if(addsearch2t.name >= 1 && addsearch2t.name <= max4table)
			addsearch2t.name++;
		else if(addsearch2t.name > max4table)
			return false;
		else
			addsearch2t.name = 1;
		var adoConn = new ActiveXObject("ADODB.Connection");
		var adoRS = new ActiveXObject("ADODB.Recordset");
		var TableName = document.getElementById("ID#"+ID+"").name;
		var fill = "<option value=0></option>";
		adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
		adoRS.Open("Select * From ["+TableName+"]", adoConn);
		for(var i = 1; i < adoRS.Fields.Count; i++)
		{
			var type = adoRS.fields(i).type;
			var name = adoRS.fields(i).name;
			fill += "<option id="+i+" name="+type+" >"+name+"</options>";
		}
		adoRS.Close();
		adoConn.Close();
		var temp = "ADDSAF#"+addsearch2t.name+"";
		addsearch2t.innerHTML = addsearch2t.innerHTML + "<div class='twelve columns'><div class='six columns'><label for='demo-name'>What Column do you want to search?</label>"+
		"<div class='select-style'><select id='ADDFS"+addsearch2t.name+"' onchange='return addsearch2text("+addsearch2t.name+")'>"+
		fill +
		"</select></div></div><div class='six columns'><label for='demo-name' >Start Typing what you want to search for</label>"+
			"<input type='hidden' id='ADDSch#"+addsearch2t.name+"' name='0192NULL' alt='"+TableName+"' value=0>"+
			"<input type='text' onkeyup='return autocompletesetupFS("+addsearch2t.name+", "+ID+")'  class='biginput' alt=0 id='ADDSAF#"+addsearch2t.name+"'>"+
				"</div></div>";
	}
	catch(e)
	{
		var e1 = "parent2child.js had an error in addsearch2t(ID) function";
		ErrorToFile(e, e1);
		location.reload();
	}
};
function addsearch2text(Fill)
{
	try{
	var select = document.getElementById("ADDFS"+Fill+"");
	var FieldName = select.options(select.selectedIndex).innerHTML;
	var HiddenF = document.getElementById("ADDSch#"+Fill+"");
	if(Fill != 1)
	{
		var r = confirm("Press OK for AND Searches.\nPress Cancel for OR Searches.");
		if (r == true) 
			HiddenF.value = "AND";
		else 
		{
			HiddenF.value = "OR";
			var TextF = document.getElementById("ADDSAF#"+Fill+"");
			TextF.style.borderColor = "#A58E8E";
		}
	}
	HiddenF.name = FieldName;
	setCookie("Search", 1, 1);
	}catch(e){ErrorToFile(e);}
};
function getsqlfromsearch()
{
	try{
	var addsearch2t = document.getElementById("addsearches2t");
	var sql = false;
	if(getCookie("Search") == 0)
		return false;
	for(var i = 1; i <= addsearch2t.name; i++)
	{
		var HiddenF1 = document.getElementById("ADDSch#"+i+"");
		var TxtBox = document.getElementById("ADDSAF#"+i+"");
		if(TxtBox.value != "")
		{
			if(sql == false)
				sql = "Select * From ["+HiddenF1.alt+"] WHERE ["+HiddenF1.name+"] LIKE '%"+TxtBox.value+"%'";
			else if(HiddenF1.value == "AND")
				sql += " AND ["+HiddenF1.name+"] LIKE '%"+TxtBox.value+"%'";
			else if(HiddenF1.value == "OR")
				sql += " OR ["+HiddenF1.name+"] LIKE '%"+TxtBox.value+"%'";
		}
	}
	return sql;
	}catch(e){ErrorToFile(e);}
}
//creats the table for fill
function generateTable(ID, Export)
{
	try{
		var TableName = document.getElementById("ID#"+ID+"").name;
		var getFill = document.getElementById("fillp4t");
		var adoConn = new ActiveXObject("ADODB.Connection");
		var adoRS = new ActiveXObject("ADODB.Recordset");
		setCookie("GTID", ID, 1);
		setCookie("TableName", TableName, 1);
		adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
		var tablecreate = "<div class='twelve columns'>"+
			"<div id='loget' class='errormessage'>" +
			"</div></div><div id='addsearches2t' ></div><div class='twelve columns'><div id='TableGBody' class='table-wrapper'><table id='TheExport' class='alt'><thead><tr>";
		if(Export == null)
			adoRS.Open("Select * From ["+TableName+"]", adoConn);
		else if(Export == false)
			adoRS.Open("Select * From ["+TableName+"]", adoConn);
		else
			adoRS.Open(Export, adoConn);
		var countsearch = 0;
		for(var i = 1; i < adoRS.Fields.Count; i++)
		{
			var name = adoRS.fields(i).name;
			var temp = adoRS.fields(i).type;
			if(temp != 203 || temp != 11)
				countsearch++;
			tablecreate += "<th>"+name+"</th>"
		
		}
		setCookie("HMS", countsearch, 1);
		tablecreate += "<th style='width:100px' >Options</th></tr></thead><tbody>";
		
			tablecreate += "<tr id = 'exporthidden' style='display: none;'>";
			for(var i = 1; i < adoRS.Fields.Count; i++)
		{
				tablecreate  += "<td>";
				tablecreate  += "<input type='radio' id='Exportxls"+i+"' name='Exportxls"+i+"' value=1 >"+
				"<label for='Exportxls"+i+"'>Export</label>";
				tablecreate  += "</br><input type='radio' id='ExportxlsN"+i+"' name='Exportxls"+i+"' >"+
				"<label for='ExportxlsN"+i+"'>Don't Export</label>";
				tablecreate  += "</td>";
		}
			tablecreate  += "<td><li onclick='return exportselected()'class='button special small'>Export Selected</li><li onclick='return selectall()'class='button special small'>Export All</li></td></tr>";
			
				
		if(!adoRS.bof)
		{
			adoRS.MoveFirst;
			while(!adoRS.eof)
			{
				tablecreate += "<tr>";
				for(var i = 1; i < adoRS.Fields.Count; i++)
				{
					var name = adoRS.fields(i).name;
					var temp = adoRS.fields(i).type;
					var value = adoRS.fields(i).value;
					
					if(temp == 203 && value == null && Export == null)
						tablecreate += "<td>No File or Memo</td>";
					else if(temp == 203 && value != null && Export == null)
					{
						if(checkiffile(TableName, name))
							tablecreate += "<td onclick='return doopen("+adoRS.fields(0)+", "+i+")'>Click to Download File<input type='hidden' id='Grab#"+adoRS.fields(0)+"N#"+i+"' name='"+name+"' value='"+value+"'/></td>";
						else
							tablecreate += "<td onclick='return makememo("+adoRS.fields(0)+", "+i+")' >Click for Memo<input type='hidden' id='Grab#"+adoRS.fields(0)+"N#"+i+"' name='"+name+"' value='"+value+"'/></td>";
					}
					else if(temp == 7)
					{	//this part is from the accessdb.js toolkit
						if(value != null)
							var tempval = (new Date((value))).format("MMM dd, yyyy");
						tablecreate += "<td>"+tempval+"</td>";
					}
					else if(temp == 6)
					{
						if(value == null)
							tablecreate += "<td></td>";
						else
							tablecreate += "<td>$"+value+"</td>";
					}
					else
					{
						if(checkiffile(TableName, name))
							tablecreate += "<td>file://"+value+"</td>";
						else
						{
							if(value != null)
								tablecreate += "<td>"+value+"</td>";
							else
								tablecreate += "<td></td>";
						}
					}
				}
				tablecreate += "<td><li onclick='return editrow("+adoRS.fields(0).value+", "+ID+")' class='button special small'>Edit</li>";
				if(deleteenabled() == true)
					tablecreate +="<li onclick='return editrow("+adoRS.fields(0).value+", "+ID+", "+ID+")' class='button special small'>Delete</li></td></tr>";
				else
					tablecreate += "</td></tr>"
				adoRS.movenext;
			}
		}
		tablecreate += "</tr></tbody></table></div></div><div id='moveexport' onclick='return showexport()' class='export'>&nbsp;&nbsp;&nbsp;</div>"+
		"<div id='addsearch' onclick='return addsearch2t("+ID+")' class='addsearch'></div>"+
		"<a id='close' href='#' class='close'>&nbsp;&nbsp;&nbsp;</a>";
		getFill.innerHTML = tablecreate;
		adoRS.Close();
		adoConn.Close();
	}
	catch(e)
	{
		var e1 = "parent2child.js had an error in generateTable(ID, Export)) function";
		ErrorToFile(e, e1);
		location.reload();
	}
};
function openfile(input)
{
	try{
	var origurl = window.location.href;
	var filename = origurl.replace("index.hta", "")
	var filename = filename.replace("adminpanel.hta", "")
	var filename = filename.replace("processviewer.hta", "")
	filename += "ImportantInformation/"+input;
	var test = window.open(filename);
	}catch(e){ErrorToFile(e);}
}
function dlatt(input)
{
	try{
	var Grab = document.getElementById("Grab#"+input+"");
	var origurl = window.location.href;
	var filename = origurl.replace("index.hta", "")
	var filename = filename.replace("adminpanel.hta", "")
	var filename = filename.replace("processviewer.hta", "")
	filename += "DatabaseFiles/"+Grab.value;
	window.open(filename);
	}catch(e){ErrorToFile(e);}
}
function doopen(ID,i)
{
	try{
	var Grab = document.getElementById("Grab#"+ID+"N#"+i+"");
	var origurl = window.location.href;
	var filename = origurl.replace("index.hta", "")
	var filename = filename.replace("adminpanel.hta", "")
	var filename = filename.replace("processviewer.hta", "")
	filename += "DatabaseFiles/"+Grab.value;
	window.open(filename);
	}catch(e){ErrorToFile(e);}
}
function checkiffile(TN, FN){
	try{
		var sql = "Select * From HyperlinkCheck Where TableName='"+TN+"' AND FieldName='"+FN+"'";
		var myDB = new ACCESSdb("Abcon.mdb");	
		if(myDB.query(sql))
			return true;
		else
			return false;
	}
	catch(e)
	{
		var e1 = "parent2child.js had an error in checkiffile(TN, FN) function";
		ErrorToFile(e, e1);
		location.reload();
	}
}
function generatefill(ID)
{
	try{
		var TableName = document.getElementById("ID#"+ID+"").name;
		var getFill = document.getElementById("fillp4f");
		getFill.innerHTML = "<div class='twelve columns$'>"+
			"<div id='loge' class='errormessage'>" +
			"</div></div>"+
			"<div class='u-full-width'>"+
			"<div class='twelve columns$'>"+
			"<label for='demo-name' class='Title4Form'>What Do You Want to Add to "+TableName+"</label>"+
			"<div class='twelve columns$'>"+
			"<div id='loge' class='errormessage'>"+
			"</div><form method='post' action=''>"+
		"<div 'twelve columns'><div id='AddFill'></div></div>"+
		"<input type='hidden' id='TxtFill'  >"+
		"<input type='hidden' id='MemoFill'  >"+
		"<input type='hidden' id='LongFill'  >"+
		"<input type='hidden' id='DouFill' >"+
		"<input type='hidden' id='CurrFill' >"+
		"<input type='hidden' id='DateFill' >"+
		"<input type='hidden' id='CheckFill' >"+
		"<input type='hidden' id='DropFill' >"+
		"<input type='hidden' id='hyperfill' >"+
		"<input type='hidden' id='Auto-Fill' >"+
		"<div id='changeoptions' class='twelve columns$'>"+
			"<ul class='actions'>"+
				"<li class='six columns'><input type='submit' onclick='return createrecord4autofill("+ID+")' value='Add Information' class='special' /></li>"+
				"<li class='six columns'><input type='reset' value='Reset Form' /></li>"+
			"</ul>"+
		"</div></div></form><a id='close' href='#' class='close'>&nbsp;&nbsp;&nbsp;</a>";
		var itsFill = document.getElementById("AddFill");
		itsFill.name = 0;
		var adoConn = new ActiveXObject("ADODB.Connection");
		var adoRS = new ActiveXObject("ADODB.Recordset");
		adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
		
		adoRS.Open("Select * From ["+TableName+"]", adoConn);
		for(var i = 1; i < adoRS.Fields.Count; i++)
		{
			var temp = adoRS.fields(i).type;
			switch(temp)
			{
				case 202:	TextFill(TableName, adoRS.fields(i).name);
							break;
				case 203:	MemoFill(TableName, adoRS.fields(i).name);
							break;
				case 3:		LongFill(adoRS.fields(i).name);
							break;
				case 5:		DoubleFill(adoRS.fields(i).name);
							break;
				case 6:		CurrFill(adoRS.fields(i).name);
							break;
				case 7:		DateFill(adoRS.fields(i).name);
							break;
				case 11:	CheckFill(adoRS.fields(i).name);
							break;
			}
		}
		adoRS.Close();
		adoConn.Close();
	}
	catch(e)
	{
		var e1 = "parent2child.js had an error in generatefill(id) function";
		ErrorToFile(e, e1);
		window.location.reload();
	}
	
};
function CheckFill(nameoffield){
	try{
		//check if drop down or if it is a normal text box.
		var itsCheck = document.getElementById("CheckFill");
		var innerfill = document.getElementById("AddFill");
		if(itsCheck.name >= 1)
				itsCheck.name++;
			else
				itsCheck.name = 1;
			innerfill.innerHTML += "<div class='twelve columns'><label for='demo-category'>"+nameoffield+"</label>"+
			"<div class='five columns'><input type='radio' id='CheckYFill"+itsCheck.name+"' name='"+nameoffield+"'>"+
			"<label for='CheckYFill"+itsCheck.name+"'>Yes</label></div>"+
			"<div class='five columns'><input type='radio' id='CheckNFill"+itsCheck.name+"' name='"+nameoffield+"' checked/>"+
			"<label for='CheckNFill"+itsCheck.name+"'>No</label></div></div>";
	}catch(e){ErrorToFile(e);}
}
function datedropb(name)
{
	try{
		 $( name).datepicker({
		  showOtherMonths: true,
		  selectOtherMonths: true,
		  dateFormat: "mm/dd/yy"
		});
	}catch(e){ErrorToFile(e);}
};
function DateFill(nameoffield){
	try{
		//check if drop down or if it is a normal text box.
		var itsDate = document.getElementById("DateFill");
		var innerfill = document.getElementById("AddFill");
		var format = "dd/mm/yy";
		if(itsDate.name >= 1)
				itsDate.name++;
			else
				itsDate.name = 1;
		var temp = "DateFill"+itsDate.name+"";
			innerfill.innerHTML +="<div class='twelve columns'>"+
			"<label for='demo-name'>"+nameoffield+"</label>"+
			"<input type='text' name='"+nameoffield+"' id='"+temp+"' onmouseover='return datedropb("+temp+")' onclick='return datedropb("+temp+")' value='MM/DD/YYYY' />"+
		"</div>";
	}catch(e){ErrorToFile(e);}
}
function CurrFill(nameoffield){
	try{//check if drop down or if it is a normal text box.
		var itsCurr = document.getElementById("CurrFill");
		var innerfill = document.getElementById("AddFill");
		if(itsCurr.name >= 1)
				itsCurr.name++;
			else
				itsCurr.name = 1;
			innerfill.innerHTML +="<div class='twelve columns'>"+
			"<label for='demo-name'>"+nameoffield+"</label>"+
			"<input type='text' name='"+nameoffield+"' id='CurrFill"+itsCurr.name+"' value='' />"+
		"</div>";
	}catch(e){ErrorToFile(e);}
}
function DoubleFill(nameoffield){
	try{//check if drop down or if it is a normal text box.
		var itsDou = document.getElementById("DouFill");
		var innerfill = document.getElementById("AddFill");
		if(itsDou.name >= 1)
				itsDou.name++;
			else
				itsDou.name = 1;
			innerfill.innerHTML +="<div class='twelve columns'>"+
			"<label for='demo-name'>"+nameoffield+"</label>"+
			"<input type='text' name='"+nameoffield+"' id='DouFill"+itsDou.name+"' value='' />"+
		"</div>";
	}catch(e){ErrorToFile(e);}
}
function MemoFill(tableName, nameoffield){
	try{
		var adoConn = new ActiveXObject("ADODB.Connection");
		var adoRS = new ActiveXObject("ADODB.Recordset");
		adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
		
		//check if drop down or if it is a normal text box.
		var itsMemo = document.getElementById("MemoFill");
		var itsHyper = document.getElementById("hyperfill");
		var innerfill = document.getElementById("AddFill");
		adoRS.Open("Select * From HyperlinkCheck WHERE TableName='"+tableName+"' AND FieldName='"+nameoffield+"'", adoConn);
		
		if(!adoRS.bof)
		{
			if(itsHyper.name >= 1)
				itsHyper.name++;
			else
				itsHyper.name = 1;
			adoRS.MoveFirst;
			innerfill.innerHTML += "<label for='demo-message'>"+nameoffield+"</label><div class='twelve columns'><label class='fileContainer'>Click Here to Upload File And Create Hyperlink <input type='file' name="+adoRS.Fields(0)+" id='HyperFill"+itsHyper.name+"' alt='"+nameoffield+"' onchange=''/></label></div>";
		}
		else
		{	
			if(itsMemo.name >= 1)
				itsMemo.name++;
			else
				itsMemo.name = 1;
			innerfill.innerHTML += "<div class='twelve columns'>"+
				"<label for='demo-message'>"+nameoffield+"</label>"+
				"<textarea name='"+nameoffield+"' id='MemoFill"+itsMemo.name+"' rows='4'></textarea>"+
			"</div>";
		}
	}
	catch(e)
	{
		var e1 = "parent2child.js had an error in MemoFill(tableName, nameoffield) function";
		ErrorToFile(e, e1);
		location.reload();
	}
}
function LongFill(nameoffield){
	try{//check if drop down or if it is a normal text box.
		var itsLong = document.getElementById("LongFill");
		if(itsLong.name >= 1)
				itsLong.name++;
			else
				itsLong.name = 1;
			innerfill.innerHTML +="<div class='twelve columns'>"+
			"<label for='demo-name'>"+nameoffield+"</label>"+
			"<input type='text' name='"+nameoffield+"' id='LongFill"+itsLong.name+"' value='' />"+
		"</div>";
	}catch(e){ErrorToFile(e);}
}
function TextFill(tableName, nameoffield){
	try{
		//check if drop down or if it is a normal text box.
		var itstxt = document.getElementById("TxtFill");
		var itsDrop = document.getElementById("DropFill");
		var itsFill = document.getElementById("Auto-Fill");
		var innerfill = document.getElementById("AddFill");
		var adoConn = new ActiveXObject("ADODB.Connection");
		var adoRS = new ActiveXObject("ADODB.Recordset");
		var myDB = new ACCESSdb("Abcon.mdb", {showErrors:true});
		adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
		var sqlc = "Select * From LookupTable WHERE RecTableName='"+tableName+"' AND RecFieldName='"+nameoffield+"' AND Type='D'"
		adoRS.Open("Select * From LookupTable WHERE RecTableName='"+tableName+"' AND RecFieldName='"+nameoffield+"'", adoConn);
		if(myDB.query(sqlc) &&  !adoRS.bof)
		{
			adoRS.MoveFirst;
			if(!adoRS.eof)
			{
				if(itsDrop.name >= 1)
					itsDrop.name++;
				else
					itsDrop.name = 1;
				var RS = new ActiveXObject("ADODB.Recordset");
				var Temp = false;
				RS.open("Select "+adoRS.fields(2)+" FROM "+adoRS.Fields(1).value+"", adoConn);
				if(!RS.bof)
				{
				//fills out the options in the selection box only if there are elements to add 
				//otherwise asks if you want to add a process
				RS.MoveFirst;
				temp = "<option value=0></option>";
			
				while(!RS.eof)
				{
					temp += "<option >"+RS.Fields(0)+"</options>";
					RS.movenext;
				}
				innerfill.innerHTML += "<div class='twelve columns'>"+
				"<label for='demo-name'>"+nameoffield+"</label>"+
				"<div class='select-style'><select id='DropFill"+itsDrop.name+"' name='"+nameoffield+"'>"+temp+
				"</select></div>";
				RS.Close();
				}
			}
		}
		else if(!myDB.query(sqlc) &&  !adoRS.bof)
		{
			if(itsFill.name >= 1)
				itsFill.name++;
			else
				itsFill.name = 1;
			innerfill.innerHTML += "<div class='twelve columns'>"+
			"<label for='demo-name'>"+nameoffield+"</label>"+
			"<p id='AutoFI#"+itsFill.name+"' class='arrhide'></p>"+
			"<input type='text' onkeyup='return autocompletesetup("+itsFill.name+")' name='"+nameoffield+"' class='biginput' id='AutoF#"+itsFill.name+"'>"+
			"</div>";
			adoRS.MoveFirst;
			
			var narray = [];
			while(!adoRS.eof)
			{
				narray.push(adoRS.fields("GetTableName").value);	//tablename
				narray.push(adoRS.fields("GetFieldName").value);	//fieldname
				adoRS.movenext;
			}
			adoRS.MoveFirst;
			narray.toString();
			document.getElementById("AutoFI#"+itsFill.name+"").innerHTML = narray;
			adoRS.Close();
			adoConn.Close();
			
		}
		else
		{
			
			if(itstxt.name >= 1)
				itstxt.name++;
			else
				itstxt.name = 1;
			innerfill.innerHTML +="<div class='twelve columns'>"+
			"<label for='demo-name'>"+nameoffield+"</label>"+
			"<input type='text' name='"+nameoffield+"' id='TxtFill"+itstxt.name+"' value='' />"+
		"</div>";
		}
	}
	catch(e)
	{
		var e1 = "parent2child.js had an error in TextFill(tableName, nameoffield) function";
		ErrorToFile(e, e1);
		location.reload();
	}
	
}
function uniq(a) {
	try{
		var prims = {"boolean":{}, "number":{}, "string":{}}, objs = [];

		return a.filter(function(item) {
			var type = typeof item;
			if(type in prims)
				return prims[type].hasOwnProperty(item) ? false : (prims[type][item] = true);
			else
				return objs.indexOf(item) >= 0 ? false : objs.push(item);
		});
	}catch(e){ErrorToFile(e);}
}
//basically grabs the autoComplete and fills it with selected info.
function autocompletesetupFS(Fill, ID)
{	
	try{
		var FieldName = document.getElementById("ADDSch#"+Fill+"").name;
		var select = document.getElementById("ADDFS"+Fill+"");
		if(FieldName == "0192NULL" || select.selectedIndex == 0)
			return false;
		var TableName = document.getElementById("ADDSch#"+Fill+"").alt;
		var type;
		if(document.getElementById("ADDSAF#"+Fill+"").alt == 0)
		{
			var adoConn = new ActiveXObject("ADODB.Connection");
			var RS = new ActiveXObject("ADODB.Recordset");
			adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
			RS.open("Select DISTINCT ["+FieldName+"] FROM ["+TableName+"]", adoConn);
			var newArray = [];
			var check = false;
			if(!RS.bof)
			{				
				RS.MoveFirst;
				while(!RS.eof)
				{
					var type = RS.Fields(0).type;
					if(type != 7)
						newArray.push(RS.Fields(0).value);
					else
					{	
						var tempval = (new Date((RS.Fields(0)))).format("M/dd/yyyy");
						newArray.push(tempval);	
					}
					RS.movenext;
				}
			}
			RS.Close();
			adoConn.Close();
			new autoComplete({
				selector: "input[id='ADDSAF#"+ parseInt(Fill)+"']",
				minChars: 1,
				source: function(term, suggest){
					term = term.toLowerCase();
					var choices = newArray;
					var suggestions = [];
					for (i=0;i<newArray.length;i++)
						if (~choices[i].toLowerCase().indexOf(term)) suggestions.push(choices[i]);
					suggest(uniq(suggestions));
					
				}
			});
			document.getElementById("ADDSAF#"+Fill+"").alt = 1;
		}
		if(type == 7)
			datedropb("ADDSAF#"+ parseInt(Fill)+"");
		var sql = getsqlfromsearch();
		if(sql == false)
		{
			return false;
		}
		var adoConn = new ActiveXObject("ADODB.Connection");
		var adoRS = new ActiveXObject("ADODB.Recordset");
		adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
		adoRS.Open(sql, adoConn);
		var tablecreate = "<table id='TheExport' class='alt'><thead><tr>";
		var gettablebody = document.getElementById("TableGBody");
		for(var i = 1; i < adoRS.Fields.Count; i++)
		{
			var name = adoRS.fields(i).name;
			var type = adoRS.fields(i).type;
			tablecreate += "<th>"+name+"</th>"
		}
		tablecreate += "<th style='width:100px' >Options</th></tr></thead><tbody>";
		tablecreate += "<tr id = 'exporthidden' style='display: none;'>";
		for(var i = 1; i < adoRS.Fields.Count; i++)
			{
				tablecreate  += "<td>";
				tablecreate  += "<input type='radio' id='Exportxls"+i+"' name='Exportxls"+i+"' value=1 >"+
				"<label for='Exportxls"+i+"'>Export</label>";
				tablecreate  += "</br><input type='radio' id='ExportxlsN"+i+"' name='Exportxls"+i+"' >"+
				"<label for='ExportxlsN"+i+"'>Don't Export</label>";
				tablecreate  += "</td>";
			}
		tablecreate  += "<td><li onclick='return exportselected()'class='button special small'>Export Selected</li><li onclick='return selectall()'class='button special small'>Export All</li></td></tr>";
		if(!adoRS.bof)
		{
			adoRS.MoveFirst;
			while(!adoRS.eof)
			{
				tablecreate += "<tr>";
				for(var i = 1; i < adoRS.Fields.Count; i++)
				{
					var name = adoRS.fields(i).name;
					var temp = adoRS.fields(i).type;
					var value = adoRS.fields(i).value;
					if(temp == 203 && value == null)
						tablecreate += "<td>No File or Memo</td>";
					else if(temp == 203 && value != null)
					{
						if(checkiffile(TableName, name))
							tablecreate += "<td onclick='return doopen("+adoRS.fields(0)+", "+i+")'>Click to Download File<input type='hidden' id='Grab#"+adoRS.fields(0)+"N#"+i+"' name='"+name+"' value='"+value+"'/></td>";
						else
							tablecreate += "<td onclick='return makememo("+adoRS.fields(0)+", "+i+")' >Click for Memo<input type='hidden' id='Grab#"+adoRS.fields(0)+"N#"+i+"' name='"+name+"' value='"+value+"'/></td>";
					}
					else if(temp == 7)
					{	//this part is from the accessdb.js toolkit
						if(value != null)
							var tempval = (new Date((value))).format("MMM dd, yyyy");
						tablecreate += "<td>"+tempval+"</td>";
					}
					else if(temp == 6)
					{
						if(value == null)
							tablecreate += "<td></td>";
						else
							tablecreate += "<td>$"+value+"</td>";
					}
					else
					{
						if(value != null)
							tablecreate += "<td>"+value+"</td>";
						else
							tablecreate += "<td></td>";
					}
				}
					tablecreate += "<td><li onclick='return editrow("+adoRS.fields(0).value+", "+ID+")' class='button special small'>Edit</li>";
					if(deleteenabled() == true)
						tablecreate +="<li onclick='return editrow("+adoRS.fields(0).value+", "+ID+", "+ID+")' class='button special small'>Delete</li></td></tr>";
					else
						tablecreate += "</td></tr>"
					adoRS.movenext;
			}
			gettablebody.innerHTML = tablecreate;
		}
	}
	catch(e)
	{
		var e1 = "parent2child.js had an error in autocompletesetupFS(Fill, ID) function";
		ErrorToFile(e, e1);
		location.reload();
	}
}
function autocompletesetup(Fill)
{	
	try{
		if(document.getElementById("AutoFI#"+Fill+"").innerHTML != 0)
		{
			var newArray = [];
			var makeautofill = document.getElementById("AutoFI#"+Fill+"").innerHTML;
			var testarray = [];
			TFarray = makeautofill.split(",");
			for(var i = 0; i < TFarray.length;i++)
			{
				var TableName = TFarray[i];
				i++;
				var FieldName = TFarray[i];
				var adoConn = new ActiveXObject("ADODB.Connection");
				var RS = new ActiveXObject("ADODB.Recordset");
				adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
				RS.open("Select DISTINCT ["+FieldName+"] FROM ["+TableName+"]", adoConn);
				var check = false;
				if(!RS.bof)
				{
					RS.MoveFirst;
					while(!RS.eof)
					{
						var type = RS.Fields(0).type;
						if(type != 7)
							newArray.push(RS.Fields(0).value);
						else
						{	
							var tempval = (new Date((RS.Fields(0)))).format("M/dd/yyyy");
							newArray.push(tempval);	
						}
						RS.movenext;
					}
				}
				RS.Close();
				adoConn.Close();
			}
		
			new autoComplete({
				selector: "input[id='AutoF#"+ parseInt(Fill)+"']",
				minChars: 1,
				source: function(term, suggest){
					term = term.toLowerCase();
					var choices = newArray;
					var suggestions = [];
					for (i=0;i<newArray.length;i++)
						if (~choices[i].toLowerCase().indexOf(term)) suggestions.push(choices[i]);
					suggest(uniq(suggestions));
					
				}
			});
			document.getElementById("AutoFI#"+Fill+"").innerHTML = 0;
		}
	}
	catch(e)
	{
		var e1 = "parent2child.js had an error in autocompletesetup(Fill) function";
		ErrorToFile(e, e1);
		location.reload();
	}
}