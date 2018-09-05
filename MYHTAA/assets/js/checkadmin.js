function backtopanel(){
	window.location = "adminpanel.hta";
}
//A Simple error message that is shared and used by multiple other functions
function titleerror()
{
	var loge = document.getElementById("loge");
	loge.innerHTML = "";
	loge.style.color = "red";
	loge.style.backgroundColor = "#fee0c6";
	loge.style.borderColor = "#7e3d3f";
	loge.innerHTML= "Invalid Title, Please have a Title." + "<br />";
}
function setcolorofloge(i)
{
	var loge = document.getElementById("loge");
	loge.innerHTML = "";
	//red
	if(i == 0)
	{
		loge.style.color = "red";
		loge.style.backgroundColor = "#fee0c6";
		loge.style.borderColor = "#7e3d3f";
	}
	//white
	else{
		loge.style.color = "White";
		loge.style.backgroundColor = "#4c6c96";
		loge.style.borderColor = "#4b2426";
	}
}
// an means add new process to the homepage page in the accessdb
function anprocess()
{
	try{
		var Title = document.getElementById("proctitle");
		var Info = document.getElementById("procinfo");
		if(Title.value == "")
		{
			Title.focus();
			titleerror();
			return false;
		}
		else
		{
			//uses accessdb.js
			var myDB = new ACCESSdb("Abcon.mdb", {showErrors:true});
			var SQL = "Insert into HomePage(PTitle, PInfo) VALUES('"+Title.value+"', '"+Info.value+"')";
			myDB.query(SQL);
			return true;
		}
	}
	catch(e)
	{
		var e1 = "checkadmin.js had an error in the anprocess() function";
		ErrorToFile(e, e1);
		location.reload();
	}
};
//edit old process in the homepage page in the accessdb
function eoprocess()
{
	try{
		var Title = document.getElementById("eproctitle");
		var Info = document.getElementById("eprocinfo");
		if(Title.value == "")
		{
			Title.focus();
			titleerror();
			return false;
		}
		else
		{
			//uses accessdb.js
			var myDB = new ACCESSdb("Abcon.mdb", {showErrors:true});
			if(Info.value != "")
				var SQL = "Update HomePage Set PTitle='"+Title.value+"', PInfo='"+Info.value+"' WHere ID="+Title.name+"";
			else
				var SQL = "Update HomePage Set PTitle='"+Title.value+"' WHERE ID="+Title.name+"";
			
			myDB.query(SQL);
		}
		return true;
	}
	catch(e)
	{
		var e1 = "checkadmin.js had an error in the eoprocess() function";
		ErrorToFile(e, e1);
		location.reload();
	}
};
//delete old process from the homepage page in the accessdb
function doprocess()
{
	try{
		var Title = document.getElementById("eproctitle");
		if(Title.value == "")
		{
			Title.focus();
			titleerror();
			return false;
		}
		else
		{
			//uses accessdb.js
			var myDB = new ACCESSdb("Abcon.mdb", {showErrors:true});
			var SQL = "Delete From HomePage where ID="+Title.name+"";
			myDB.query(SQL);
		}
		return true;
	}
	catch(e)
	{
		var e1 = "checkadmin.js had an error in the doprocess() function";
		ErrorToFile(e, e1);
		location.reload();
	}
};
//Hide old process by changing hiddens value homepage page in the accessdb
function hoprocess()
{
	try{
		var Title = document.getElementById("eproctitle");
		if(Title.value == "")
		{
			Title.focus();
			titleerror();
		}
		else
		{
			//uses accessdb.js
			var myDB = new ACCESSdb("Abcon.mdb", {showErrors:true});
			var SQL = "Select * From HomePage where ID="+Title.name+" AND Hidden='F'";
			if(myDB.query(SQL))
			{
				//sets hidden to true
				SQL = "Update HomePage Set Hidden='T' WHere ID="+Title.name+"";
				myDB.query(SQL);
			}
			else
			{
				//sets hiddent to false
				SQL = "Update HomePage Set Hidden='F' WHere ID="+Title.name+"";
				myDB.query(SQL);
			}
		}
	}
	catch(e)
	{
		var e1 = "checkadmin.js had an error in the hoprocess() function";
		ErrorToFile(e, e1);
		location.reload();
	}
};
//The query to add and delete autofills
function AutoFillQuery(type)
{
	try{
		var filldrop = document.getElementById("AddTTN");
		var TTN = filldrop.options[filldrop.selectedIndex].text
		var filldrop = document.getElementById("AddTFN");
		var TFN = filldrop.options[filldrop.selectedIndex].text
		var filldrop = document.getElementById("AddFTN");
		var FTN = filldrop.options[filldrop.selectedIndex].text
		var filldrop = document.getElementById("AddFFN");
		var FFN = filldrop.options[filldrop.selectedIndex].text
		var myDB = new ACCESSdb("Abcon.mdb", {showErrors:true});
		if(type == 1)
			myDB.query("Insert into Lookuptable(GetTableName, GetFieldName, RecTableName, RecFieldName, Type) VALUES('"+FTN+"','"+FFN+"','"+TTN+"','"+TFN+"','A')")
		else
			myDB.query("Delete from Lookuptable Where GetTableName='"+FTN+"' AND GetFieldName='"+FFN+"' AND RecTableName='"+TTN+"' AND RecFieldName ='"+TFN+"'")
	}
	catch(e)
	{
		var e1 = "checkadmin.js had an error in the AutoFillQuery(type) function";
		ErrorToFile(e, e1);
		location.reload();
	}
}
//this one doesnt use accessdb.js I needed more fucntionality so I did it myself
//this creats a drop down list of the selected tables.
function dropdownmeu(p1, option)
{
	try{
		if(option == null)
		{
			var getdd = document.getElementById("SelTDrop"+p1+"");
			var filldrop = document.getElementById("SelFDrop"+p1+"");
		}
		else if(option == 1)
		{
			var getdd = document.getElementById("SelATDrop"+p1+"");
			var filldrop = document.getElementById("SelAFDrop"+p1+"");
		}
		else if(option == 2)
		{
			var getdd = document.getElementById("AddTTN");
			var filldrop = document.getElementById("AddTFN");
		}
		else if(option == 3)
		{
			var getdd = document.getElementById("AddFTN");
			var filldrop = document.getElementById("AddFFN");
		}
		var tableName = getdd.options[getdd.selectedIndex].text;
		var adoConn = new ActiveXObject("ADODB.Connection");
		var adoRS = new ActiveXObject("ADODB.Recordset");
		adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
	adoRS.Open("Select * From ["+tableName+"]", adoConn);
	
		//fills out the opations in the selection box only if there are elements to add 
		//otherwise asks if you want to add a process
		filldrop.options.length = 0;
		var optn = document.createElement("OPTION");
		optn.text = "";
		filldrop.options.add(optn)
		
		for(var i = 0; i < adoRS.Fields.Count; i++)
		{
			optn = document.createElement("OPTION");
			optn.text = (adoRS.fields(i).name.split(' ').join('_'));
			optn.value = (adoRS.fields(i).name.split(' ').join('_'));
			filldrop.options.add(optn)
		}
		adoRS.Close();
		adoConn.Close();
	}
	catch(e)
	{
		var e1 = "checkadmin.js had an error in the dropdownmeu(p1, option) function";
		ErrorToFile(e, e1);
		location.reload();
	}
}
//checks if we can add new columns to a table and adds it quickly
function checkANC()
{
	//
	//Start of Try Function I wont tab here
	//
	try{
	//this function calls all the checks required between the fields checking whether they are empty or have information
	// then checks that information to see if another field has the same information if that happens
	// the user gets an error message telling them to make sure all the field names are different if they are not blank.
	// if they are blank they are not created when the table is created.
	var bool = testallcases4CNT();
	if(bool == true)
	{
		var loge = document.getElementById("loge");
		setcolorofloge(0);
		var getTN = document.getElementById("OTableName");
		var tableName = stringfix(getTN.options[getTN.selectedIndex].text);
		var error  = "Some Symbols do Not Work AS Column Names(Please View User Manuals for More Information)\n";
		var myDB = new ACCESSdb("Abcon.mdb", {showErrors:true});
		var counter = document.getElementById("texta").value;
		var totalcnt = document.getElementById("totccnt");
		var txti = 1, memoi = 1, inti = 1, doui = 1, curri = 1, datei = 1, checki = 1, dropi = 1,hyperi = 1, afilli = 1;
		for(var i = 1; i <= totalcnt.value; i++)
		{
			//get the text boxes in the table
			var getlocation = document.getElementById("addbox0"+i+"");
			if(getlocation.title == "Text")
			{
				txttemp1 = document.getElementById("DataText"+txti+"");
				var SQL = "ALTER TABLE ["+tableName+"] ADD COLUMN ["+txttemp1.value+"] varchar(255)";
				myDB.query(SQL);
				var sqldd  = "Insert Into Lookuptable(GetTableName, GetFieldName, RecTableName, RecFieldName, Type) VALUES('"+tableName+"','"+txttemp1.value+"','"+tableName+"','"+txttemp1.value+"', 'A')";
				myDB.query(sqldd);
				txti++;
			}
			//get the memo boxes in the table
			else if(getlocation.title == "Memo")
			{
				txttemp1 = document.getElementById("DataMemo"+memoi+"");
				var SQL = "ALTER TABLE ["+tableName+"] ADD COLUMN ["+txttemp1.value+"] MEMO";
				myDB.query(SQL);
				memoi++;
			}
			//get the integer boxes in the table
			else if(getlocation.title == "Int")
			{
				txttemp1 = document.getElementById("DataInt"+inti+"");
				var SQL = "ALTER TABLE ["+tableName+"] ADD COLUMN ["+txttemp1.value+"] Long";
				myDB.query(SQL);
				inti++;
			}
			//get the double boxes in the table
			else if(getlocation.title == "Double")
			{
				txttemp1 = document.getElementById("DataDouble"+doui+"");
				var SQL = "ALTER TABLE ["+tableName+"] ADD COLUMN ["+txttemp1.value+"] Double";
				myDB.query(SQL);
				doui++;
			}
			//get the currency boxes in the table
			else if(getlocation.title == "Curr")
			{
				txttemp1 = document.getElementById("DataCurr"+curri+"");
				var SQL = "ALTER TABLE ["+tableName+"] ADD COLUMN ["+txttemp1.value+"] Currency";
				myDB.query(SQL);
				curri++;
			}
			//get the Date boxes in the table
			else if(getlocation.title == "Date")
			{
				txttemp1 = document.getElementById("DataDate"+datei+"");
				var SQL = "ALTER TABLE ["+tableName+"] ADD COLUMN ["+txttemp1.value+"] DATETIME";
				myDB.query(SQL);
				datei++;
			}
			//get the Check boxes in the table
			else if(getlocation.title == "Check")
			{
				txttemp1 = document.getElementById("DataCheck"+checki+"");
				var SQL = "ALTER TABLE ["+tableName+"] ADD COLUMN ["+txttemp1.value+"] YESNO";
				myDB.query(SQL);
				checki++;
			}
			//get the DROPDOWN boxes in the table
			else if(getlocation.title == "Drop")
			{
				var txttemp1 = document.getElementById("SDropName"+dropi+"");
				var getdd = document.getElementById("SelTDrop"+dropi+"");
				
				var filldrop = document.getElementById("SelFDrop"+dropi+"");
				
				if(filldrop.options[filldrop.selectedIndex].text == "")
				{
					loge.innerHTML = "Please Select the Column you want to Populate the Drop-Down with.<br/>";
					return false;
				}
				if(getdd.options[getdd.selectedIndex].text == "")
				{
					loge.innerHTML = "Please Select the Table you want to get the Columns For.<br/>";
					return false;
				}
				var SQL = "ALTER TABLE ["+tableName+"] ADD COLUMN ["+txttemp1.value+"] varchar(255)";
				myDB.query(SQL);
				var sqldd  = "Insert Into Lookuptable(GetTableName, GetFieldName, RecTableName, RecFieldName, Type) VALUES('"+getdd.options[getdd.selectedIndex].text+"','"+filldrop.options[filldrop.selectedIndex].text+"','"+tableName+"','"+txttemp1.value+"', 'D')";
				myDB.query(sqldd);
				dropi++;
			}
			//for hyperlink boxes
			else if(getlocation.title == "Hyper")
			{
				txttemp1 = document.getElementById("DataHyper"+hyperi+"");
				var SQL = "ALTER TABLE ["+tableName+"] ADD COLUMN ["+txttemp1.value+"] MEMO";
				myDB.query(SQL);
				var sqldd  = "Insert Into HyperlinkCheck(TableName, FieldName) VALUES('"+tableName+"','"+txttemp1.value+"')";
				myDB.query(sqldd);
				hyperi++;
			}
			//for autofilled boxes
			else if(getlocation.title == "AFill")
			{
				var txttemp1 = document.getElementById("AutoName"+afilli+"");
				var getdd = document.getElementById("SelATDrop"+afilli+"");
				
				var filldrop = document.getElementById("SelAFDrop"+afilli+"");
				
				if(filldrop.options[filldrop.selectedIndex].text == "")
				{
					loge.innerHTML = "Please Select the Column you want to Populate the Drop-Down with.<br/>";
					return false;
				}
				if(getdd.options[getdd.selectedIndex].text == "")
				{
					loge.innerHTML = "Please Select the Table you want to get the Columns For.<br/>";
					return false;
				}
				var SQL = "ALTER TABLE ["+tableName+"] ADD COLUMN ["+txttemp1.value+"] varchar(255)";
				myDB.query(SQL);
				var sqldd  = "Insert Into Lookuptable(GetTableName, GetFieldName, RecTableName, RecFieldName, Type) VALUES('"+getdd.options[getdd.selectedIndex].text+"','"+filldrop.options[filldrop.selectedIndex].text+"','"+tableName+"','"+txttemp1.value+"', 'A')";
				myDB.query(sqldd);
				var sqldd  = "Insert Into Lookuptable(GetTableName, GetFieldName, RecTableName, RecFieldName, Type) VALUES('"+tableName+"','"+txttemp1.value+"','"+tableName+"','"+txttemp1.value+"', 'A')";
				myDB.query(sqldd);
				afilli++;
			}
		}
		setcolorofloge(1);
		loge.innerHTML = "The Column(s) has Been Added.";
	}
	//
	//End of Try Start of Catch
	//
	}
	catch(e)
	{
		var e1 = "checkadmin.js had an error in checkANC() function";
		ErrorToFile(e, e1);
		location.reload();
	}
}
//checks if we can create a new table witht he inputed values then creates a table with those values/.
function checkCNT()
{
	//
	//Start of Try Function I wont tab here
	//
	try{
	//this function calls all the checks required between the fields checking whether they are empty or have information
	// then checks that information to see if another field has the same information if that happens
	// the user gets an error message telling them to make sure all the field names are different if they are not blank.
	// if they are blank they are not created when the table is created.
	var bool = testallcases4CNT();
	if(bool == true)
	{
		var loge = document.getElementById("loge");
		setcolorofloge(0);
		var tableName = document.getElementById("tableName");
		var test4add = false;
		var dropdowninsert = false;
		var sql = "Create Table ["+tableName.value+"](";
		sql += "ID AUTOINCREMENT NOT NULL PRIMARY KEY,";
		var totalcnt = document.getElementById("totccnt");
		var sqladd2lu = [];
		var txti = 1, memoi = 1, inti = 1, doui = 1, curri = 1, datei = 1, checki = 1, dropi = 1,hyperi = 1, afilli = 1;
		for(var i = 1; i <= totalcnt.value; i++)
		{
			//get the text boxes in the table
			var getlocation = document.getElementById("addbox0"+i+"");
			if(getlocation.title == "Text")
			{
				txttemp1 = document.getElementById("DataText"+txti+"");
				var sqldd  = "Insert Into Lookuptable(GetTableName, GetFieldName, RecTableName, RecFieldName, Type) VALUES('"+tableName.value+"','"+txttemp1.value+"','"+tableName.value+"','"+txttemp1.value+"', 'A')";
				sqladd2lu.push(sqldd);
				if(i == 1)
					sql += "["+txttemp1.value+"] varchar(255)";
				else
					sql += ", ["+txttemp1.value+"] varchar(255)";
				txti++;
				test4add = true;
			}
			//get the memo boxes in the table
			else if(getlocation.title == "Memo")
			{
				txttemp1 = document.getElementById("DataMemo"+memoi+"");
				if(i == 1)
					sql += "["+txttemp1.value+"] Memo";
				else
					sql += ", ["+txttemp1.value+"] Memo";
				memoi++;
				test4add = true;
			}
			//get the integer boxes in the table
			else if(getlocation.title == "Int")
			{
				txttemp1 = document.getElementById("DataInt"+inti+"");
				if(i == 1)
					sql += "["+txttemp1.value+"] Long";
				else
					sql += ", ["+txttemp1.value+"] Long";
				inti++;
				test4add = true;
			}
			//get the double boxes in the table
			else if(getlocation.title == "Double")
			{
				txttemp1 = document.getElementById("DataDouble"+doui+"");
				if(i == 1)
					sql += "["+txttemp1.value+"] Double";
				else
					sql += ", ["+txttemp1.value+"] Double";
				doui++;
				test4add = true;
			}
			//get the currency boxes in the table
			
			else if(getlocation.title == "Curr")
			{
				txttemp1 = document.getElementById("DataCurr"+curri+"");
				if(i == 1)
					sql += "["+txttemp1.value+"] Currency";
				else
					sql += ", ["+txttemp1.value+"] Currency";
				curri++;
				test4add = true;
			}
			//get the Date boxes in the table
			else if(getlocation.title == "Date")
			{
				txttemp1 = document.getElementById("DataDate"+datei+"");
				if(i == 1)
					sql += "["+txttemp1.value+"] DATETIME";
				else
					sql += ", ["+txttemp1.value+"] DATETIME";
				datei++;
				test4add = true;
			}
			//get the Check boxes in the table
			else if(getlocation.title == "Check")
			{
				txttemp1 = document.getElementById("DataCheck"+checki+"");
				if(i == 1)
					sql += "["+txttemp1.value+"] YESNO";
				else
					sql += ", ["+txttemp1.value+"] YESNO";
				checki++;
				test4add = true;
			}
			//get the DROPDOWN boxes in the table
			else if(getlocation.title == "Drop")
			{
				var txttemp1 = document.getElementById("SDropName"+dropi+"");
				var getdd = document.getElementById("SelTDrop"+dropi+"");
				
				var filldrop = document.getElementById("SelFDrop"+dropi+"");
				
				if(filldrop.options[filldrop.selectedIndex].text == "")
				{
					loge.innerHTML = "Please Select the Column you want to Populate the Drop-Down with.<br/>";
					return false;
				}
				if(getdd.options[getdd.selectedIndex].text == "")
				{
					loge.innerHTML = "Please Select the Table you want to get the Columns For.<br/>";
					return false;
				}
				var sqldd  = "Insert Into Lookuptable(GetTableName, GetFieldName, RecTableName, RecFieldName, Type) VALUES('"+getdd.options[getdd.selectedIndex].text+"','"+filldrop.options[filldrop.selectedIndex].text+"','"+(tableName.value)+"','"+txttemp1.value+"', 'D')";
				sqladd2lu.push(sqldd);
				if(i == 1)
					sql += "["+txttemp1.value+"] varchar(255)";
				else
					sql += ", ["+txttemp1.value+"] varchar(255)";
				dropi++;
				test4add = true;
			}
			//for hyperlink boxes
			else if(getlocation.title == "Hyper")
			{
				txttemp1 = document.getElementById("DataHyper"+hyperi+"");
				var sqldd  = "Insert Into HyperlinkCheck(TableName, FieldName) VALUES('"+tableName.value+"','"+txttemp1.value+"')";
				sqladd2lu.push(sqldd);
				if(i == 1)
					sql += "["+txttemp1.value+"] MEMO";
				else
					sql += ", ["+txttemp1.value+"] MEMO";
				hyperi++;
				test4add = true;
			}
			//for autofilled boxes
			else if(getlocation.title == "AFill")
			{
				var txttemp1 = document.getElementById("AutoName"+afilli+"");
				var getdd = document.getElementById("SelATDrop"+afilli+"");
				
				var filldrop = document.getElementById("SelAFDrop"+afilli+"");
				
				if(filldrop.options[filldrop.selectedIndex].text == "")
				{
					loge.innerHTML = "Please Select the Column you want to Populate the Drop-Down with.<br/>";
					return false;
				}
				if(getdd.options[getdd.selectedIndex].text == "")
				{
					loge.innerHTML = "Please Select the Table you want to get the Columns For.<br/>";
					return false;
				}
				var sqldd  = "Insert Into Lookuptable(GetTableName, GetFieldName, RecTableName, RecFieldName, Type) VALUES('"+getdd.options[getdd.selectedIndex].text+"','"+filldrop.options[filldrop.selectedIndex].text+"','"+(stringfix(tableName.value))+"','"+txttemp1.value+"', 'A')";
				sqladd2lu.push(sqldd);
				var sqldd  = "Insert Into Lookuptable(GetTableName, GetFieldName, RecTableName, RecFieldName, Type) VALUES('"+tableName.value+"','"+txttemp1.value+"','"+tableName.value+"','"+txttemp1.value+"', 'A')";
				sqladd2lu.push(sqldd);
				if(i == 1)
					sql += "["+txttemp1.value+"] varchar(255)";
				else
					sql += ", ["+txttemp1.value+"] varchar(255)";
				afilli++;
				test4add = true;
			}
			
		}
		sql += ");";
		var myDB = new ACCESSdb("Abcon.mdb", {showErrors:true});
		sqltablelist = "Insert Into TableList(TableName, ParentProcessID) VALUES('"+tableName.value+"','"+(tableName.name)+"')";
		if(test4add == true)
		{
			myDB.query(sql);
			myDB.query(sqltablelist);
			for(var i = 0; i < sqladd2lu.length;i++)
				myDB.query(sqladd2lu[i]);
		}
		setcolorofloge(1);
		loge.innerHTML += "The Table has been Created.";
	}
	return true;	
	//
	//End of Try Start of Catch
	//
	}
	catch(e)
	{
		var e1 = "checkadmin.js had an error in checkCNT() function";
		ErrorToFile(e, e1);
		window.location.reload();
	}
}
//changes the table name of a table by pretty much creating a new table with the required name
//Sadly msaccess has limitations in what can be done.
// then afterwards all attached records are changed or deleted.
function ChangeTN()
{
	try{
		var TableName = document.getElementById("tableName");
		if(TableName.value != "")
		{
			var myDB = new ACCESSdb("Abcon.mdb", {showErrors:true});
			var SQL = "SELECT * INTO ["+(stringfix(TableName.value))+"] FROM ["+TableName.name+"]";
			myDB.query(SQL);
			var SQL = "Drop Table ["+(stringfix(TableName.name))+"]";
			myDB.query(SQL);
			var SQL = "Update TableList Set TableName='"+(stringfix(TableName.value))+"' Where TableName='"+TableName.name+"'";
			myDB.query(SQL);
			var SQL = "UPDATE HyperlinkCheck Set TableName='"+stringfix(TableName.value)+"' where TableName='"+TableName.name+"'";
			myDB.query(SQL);
			var SQL = "Update Lookuptable Set GetTableName='"+(stringfix(TableName.value))+"' Where GetTableName='"+TableName.name+"'";
			myDB.query(SQL);
			var SQL = "Update Lookuptable Set RecTableName='"+(stringfix(TableName.value))+"' Where RecTableName='"+TableName.name+"'";
			myDB.query(SQL);
			return true;
		}
		return false;
	}
	catch(e)
	{
		var e1 = "checkadmin.js had an error in ChangeTN() function";
		ErrorToFile(e, e1);
		window.location.reload();
	}
}
//deletes a table and all associated data
function DeleteTN()
{
	try{
		var TableName = document.getElementById("tableName");
		if(TableName.name != "")
		{
			var myDB = new ACCESSdb("Abcon.mdb", {showErrors:true});
			var SQL = "Drop Table ["+TableName.name+"]";
			myDB.query(SQL);
			var SQL = "Delete From TableList where TableName='"+TableName.name+"'";
			myDB.query(SQL);
			var SQL = "Delete From HyperlinkCheck where TableName='"+TableName.name+"'";
			myDB.query(SQL);
			var SQL = "Delete From Lookuptable where GetTableName='"+TableName.name+"'";
			myDB.query(SQL);
			var SQL = "Delete From Lookuptable where RecTableName='"+TableName.name+"'";
			myDB.query(SQL);
			return true;
		}	
		return false;
	}
	catch(e)
	{
		var e1 = "checkadmin.js had an error in DeleteTN() function";
		ErrorToFile(e, e1);
		window.location.reload();
	}
}
//Keeps the table in the database but it no longer shows up when in the application
function DiscontinueTN()
{
	try{
		var TableName = document.getElementById("tableName");
		if(TableName.name != "")
		{
			var myDB = new ACCESSdb("Abcon.mdb", {showErrors:true});
			var SQL = "Select * From TableList where TableName='"+TableName.name+"' AND Obsolete='F'";
			if(myDB.query(SQL))
			{
				SQL = "Update TableList Set Obsolete='T' WHere TableName='"+TableName.name+"'";
				myDB.query(SQL);
			}
			else
			{
				SQL = "Update TableList Set Obsolete='F' WHere TableName='"+TableName.name+"'";
				myDB.query(SQL);
			}
			return true;
		}
		return false;	
	}
	catch(e)
	{
		var e1 = "checkadmin.js had an error in DiscontinueTN() function";
		ErrorToFile(e, e1);
		window.location.reload();
	}
}
//links a table to another process
function TableLinking()
{
	try{
		var loge = document.getElementById("loge");
		var TableName = document.getElementById("tableName");
		var temp = document.getElementById("linkto")
		if(TableName.name != "")
		{
			var myDB = new ACCESSdb("Abcon.mdb", {showErrors:true});
			var SQL = "Insert into TableList(TableName, ParentProcessID) VALUES('"+(TableName.name)+"', '"+temp.options[temp.selectedIndex].id+"')";
			myDB.query(SQL);
			return true;
		}
		return false;	
	}
	catch(e)
	{
		var e1 = "checkadmin.js had an error in TableLinking() function";
		ErrorToFile(e, e1);
		window.location.reload();
	}
}
//adds an admin to the admin table
function AddAdmin()
{
	try{
		var loge = document.getElementById("loge");
		var NewAdmin = document.getElementById("NAdName")
		if(NewAdmin.value != "")
		{
			var myDB = new ACCESSdb("Abcon.mdb", {showErrors:true});
			var SQL = "Insert into AdminTable(AdminName) VALUES('"+NewAdmin.value+"')";
			myDB.query(SQL);
			return true;
		}
		return false;
	}
	catch(e)
	{
		var e1 = "checkadmin.js had an error in AddAdmin() function";
		ErrorToFile(e, e1);
		window.location.reload();
	}
}
	//this function calls all the checks required between the fields checking whether they are empty or have information
	// then checks that information to see if another field has the same information if that happens
	// the user gets an error message telling them to make sure all the field names are different if they are not blank.
	// if they are blank they are not created when the table is created.
function testallcases4CNT()
{
	
	//
	//Start of Try Function I wont tab here
	//
	try{
	var loge = document.getElementById("loge");
	setcolorofloge(0)
	var txttemp1 = document.getElementById("tableName");
	if(stringcheck(txttemp1.value) == false)
	{	
		loge.innerHTML = error;
		return false;
	}
	txttemp1.value = stringfix(txttemp1.value);
	var error = "Some symbols don't work In names(Please View Administrator Manual for More Information)";
	var bool = true;
	var test = true;
	var re = /^[0-9]/;
	//check texta to all the values in the other fields
	var counter = document.getElementById("texta").value;

	for(var i = 1 ; i <= counter && counter != null; i++)
	{
		
		txttemp1 = document.getElementById("DataText"+i+"");
		counter = document.getElementById("autofa").value;
		//checks text fields against memo fields
		for(var j = 1; j <= counter; j++)
		{
			txttemp2 = document.getElementById("AutoName"+j+"");
			if(txttemp1.value == txttemp2.value  && txttemp2.value != "")
			{
				if(bool == true)
				{
					loge.innerHTML = "Please Have Separate Names for all of your Fields.<br/>Please Look at the fields outlined in Red.";
					txttemp2.focus();
				}
				bool = false;
				txttemp2.style.borderColor = "#7e3d3f";
			}
			else
				if(txttemp2.style.borderColor != "#7e3d3f")
					txttemp2.style.backgroundColor = "#476A34";
		}
		counter = document.getElementById("hypera").value;
		//checks text fields against memo fields
		for(var j = 1; j <= counter; j++)
		{
			txttemp2 = document.getElementById("DataHyper"+j+"");
			if(txttemp1.value == txttemp2.value  && txttemp2.value != "")
			{
				if(bool == true)
				{
					loge.innerHTML = "Please Have Separate Names for all of your Fields.<br/>Please Look at the fields outlined in Red.";
					txttemp2.focus();
				}
				bool = false;
				txttemp2.style.borderColor = "#7e3d3f";
			}
			else
				if(txttemp2.style.borderColor != "#7e3d3f")
					txttemp2.style.backgroundColor = "#476A34";
		}
		counter = document.getElementById("memoa").value;
		//checks text fields against memo fields
		for(var j = 1; j <= counter; j++)
		{
			txttemp2 = document.getElementById("DataMemo"+j+"");
			if(txttemp1.value == txttemp2.value  && txttemp2.value != "")
			{
				if(bool == true)
				{
					loge.innerHTML = "Please Have Separate Names for all of your Fields.<br/>Please Look at the fields outlined in Red.";
					txttemp2.focus();
				}
				bool = false;
				txttemp2.style.borderColor = "#7e3d3f";
			}
			else
				if(txttemp2.style.borderColor != "#7e3d3f")
					txttemp2.style.backgroundColor = "#476A34";
		}
		counter = document.getElementById("inta").value;
		//checks text fields against integer fields
		for(var j = 1; j <= counter; j++)
		{
			txttemp2 = document.getElementById("DataInt"+j+"");
			if(txttemp1.value == txttemp2.value  && txttemp2.value != "")
			{
				if(bool == true)
				{
					loge.innerHTML = "Please Have Separate Names for all of your Fields.<br/>Please Look at the fields outlined in Red.";
					txttemp2.focus();
				}
				bool = false;
				txttemp2.style.borderColor = "#7e3d3f";
			}
			else
				if(txttemp2.style.borderColor != "#7e3d3f")
					txttemp2.style.backgroundColor = "#476A34";
		}
		//checks text fields against double fields
		
		counter = document.getElementById("doublea").value;
		for(var j = 1; j <= counter; j++)
		{
			txttemp2 = document.getElementById("DataDouble"+j+"");
			if(txttemp1.value == txttemp2.value && txttemp2.value != "")
			{
				if(bool == true)
				{
					loge.innerHTML = "Please Have Separate Names for all of your Fields.<br/>Please Look at the fields outlined in Red.";
					txttemp2.focus();
				}
				bool = false;
				txttemp2.style.borderColor = "#7e3d3f";
			}
			else
				if(txttemp2.style.borderColor != "#7e3d3f")
					txttemp2.style.backgroundColor = "#476A34";
		}
		//checks text fields against currency fields
		counter = document.getElementById("curra").value;
		for(var j = 1; j <= counter; j++)
		{
			txttemp2 = document.getElementById("DataCurr"+j+"");
			if(txttemp1.value == txttemp2.value && txttemp2.value != "")
			{
				if(bool == true)
				{
					loge.innerHTML = "Please Have Separate Names for all of your Fields.<br/>Please Look at the fields outlined in Red.";
					txttemp2.focus();
				}
				bool = false;
				txttemp2.style.borderColor = "#7e3d3f";
			}
			else
				if(txttemp2.style.borderColor != "#7e3d3f")
					txttemp2.style.backgroundColor = "#476A34";
		}
		//checks text fields against date fields
		counter = document.getElementById("datea").value;
		for(var j = 1; j <= counter; j++)
		{
			txttemp2 = document.getElementById("DataDate"+j+"");
			if(txttemp1.value == txttemp2.value && txttemp2.value != "")
			{
				if(bool == true)
				{
					loge.innerHTML = "Please Have Separate Names for all of your Fields.<br/>Please Look at the fields outlined in Red.";
					txttemp2.focus();
				}
				bool = false;
				txttemp2.style.borderColor = "#7e3d3f";
			}
			else
				if(txttemp2.style.borderColor != "#7e3d3f")
					txttemp2.style.backgroundColor = "#476A34";
		}
		//checks text fields against check boxes
		counter = document.getElementById("checka").value;
		for(var j = 1; j <= counter; j++)
		{
			txttemp2 = document.getElementById("DataCheck"+j+"");
			if(txttemp1.value == txttemp2.value && txttemp2.value != "")
			{
				if(bool == true)
				{
					loge.innerHTML = "Please Have Separate Names for all of your Fields.<br/>Please Look at the fields outlined in Red.";
					txttemp2.focus();
				}
				bool = false;
				txttemp2.style.borderColor = "#7e3d3f";
			}
			else
				if(txttemp2.style.borderColor != "#7e3d3f")
					txttemp2.style.backgroundColor = "#476A34";
		}
		//checks text fields against drop down boxes
		counter = document.getElementById("dropa").value;
		for(var j = 1; j <= counter; j++)
		{
			txttemp2 = document.getElementById("SDropName"+j+"");
			if(txttemp1.value == txttemp2.value && txttemp2.value != "")
			{
				if(bool == true)
				{
					loge.innerHTML = "Please Have Separate Names for all of your Fields.<br/>Please Look at the fields outlined in Red.";
					txttemp2.focus();
				}
				bool = false;
				txttemp2.style.borderColor = "#7e3d3f";
			}
			else
				if(txttemp2.style.borderColor != "#7e3d3f")
					txttemp2.style.backgroundColor = "#476A34";
		}
	}
	//checkstext fields against text fields
	var counter = document.getElementById("texta").value;
	for(var i = 1; i <= counter; i++)
	{
		txttemp1 = document.getElementById("DataText"+i+"");
		if(txttemp1.value == "")
		{
			if(bool == true)
			{
				loge.innerHTML = "Please Make Sure all your Fields have a Name.";
				txttemp1.focus();
			}
			bool = false;
			txttemp1.style.borderColor = "#7e3d3f";
		}
		else
		{
			if(stringcheck(txttemp1.value) == false)
			{
				loge.innerHTML = error;
				txttemp1.value = stringfix(txttemp1.value);
				bool = false;
			}
			else
				if(txttemp1.style.borderColor != "#7e3d3f")
						txttemp1.style.backgroundColor = "#476A34";	
			
			txttemp1.value = stringfix(txttemp1.value);
		}
		for(var j = 1; j <= counter; j++)
		{
			if(i == j && j <= counter)
				j++;
			if(j > counter)
				break;
			txttemp2 = document.getElementById("DataText"+j+"");
			if(txttemp1.value == txttemp2.value && txttemp2.value != "")
			{
				txttemp2.style.borderColor = "#7e3d3f";
				if(bool == true)
				{
					loge.innerHTML = "Please Have Separate Names for all of your Fields.<br/>Please Look at the fields outlined in Red.";
					txttemp2.focus();
				}
				bool = false;
			}
			else
				if(txttemp2.style.borderColor != "#7e3d3f")
					txttemp2.style.backgroundColor = "#476A34";
		}
	}
	//checks memo fields against memo fields
	var counter = document.getElementById("memoa").value;
	for(var i = 1; i <= counter; i++)
	{
		txttemp1 = document.getElementById("DataMemo"+i+"");
		txttemp1.value = stringfix(txttemp1.value);
		if(txttemp1.value == "")
		{
			if(bool == true)
			{
				loge.innerHTML = "Please Make Sure all your Fields have a Name.";
				txttemp1.focus();
			}
			bool = false;
			txttemp1.style.borderColor = "#7e3d3f";
		}
		else
		{
			if(stringcheck(txttemp1.value) == false)
			{
				loge.innerHTML = error;
				txttemp1.value = stringfix(txttemp1.value);
				bool = false;
			}
			else
				if(txttemp1.style.borderColor != "#7e3d3f")
						txttemp1.style.backgroundColor = "#476A34";	
			
			txttemp1.value = stringfix(txttemp1.value);
		}
		for(var j = 1; j <= counter; j++)
		{
			if(i == j && j <= counter)
				j++;	
			if(j > counter)
				break;
			txttemp2 = document.getElementById("DataMemo"+j+"");
			if(txttemp1.value == txttemp2.value && txttemp2.value != "")
			{
				txttemp2.style.borderColor = "#7e3d3f";
				if(bool == true)
				{
					loge.innerHTML = "Please Have Separate Names for all of your Fields.<br/>Please Look at the fields outlined in Red.";
					txttemp2.focus();
				}
				bool = false;
			}
			else
				if(txttemp2.style.borderColor != "#7e3d3f")
					txttemp2.style.backgroundColor = "#476A34";
		}
	}
	//checks int fields agaisnt int fields
	var counter = document.getElementById("inta").value;
	for(var i = 1; i <= counter; i++)
	{
		txttemp1 = document.getElementById("DataInt"+i+"");
		txttemp1.value = stringfix(txttemp1.value);
		if(txttemp1.value == "")
		{
			if(bool == true)
			{
				loge.innerHTML = "Please Make Sure all your Fields have a Name.";
				txttemp1.focus();
			}
			bool = false;
			txttemp1.style.borderColor = "#7e3d3f";
		}
		else
		{
			if(stringcheck(txttemp1.value) == false)
			{
				loge.innerHTML = error;
				txttemp1.value = stringfix(txttemp1.value);
				bool = false;
			}
			else
				if(txttemp1.style.borderColor != "#7e3d3f")
						txttemp1.style.backgroundColor = "#476A34";	
			
			txttemp1.value = stringfix(txttemp1.value);
		}
		for(var j = 1; j <= counter; j++)
		{
			if(i == j && j <= counter)
				j++;	
			if(j > counter)
				break;
			txttemp2 = document.getElementById("DataInt"+j+"");
			if(txttemp1.value == txttemp2.value && txttemp2.value != "")
			{
				txttemp2.style.borderColor = "#7e3d3f";				
				if(bool == true)
				{
					loge.innerHTML = "Please Have Separate Names for all of your Fields.<br/>Please Look at the fields outlined in Red.";
					txttemp2.focus();
				}
				bool = false;
				txttemp2.style.borderColor = "#7e3d3f";
			}
			else
				if(txttemp2.style.borderColor != "#7e3d3f")
					txttemp2.style.backgroundColor = "#476A34";
		}
	}
	//checks double fields against double fields
	var counter = document.getElementById("doublea").value;
	for(var i = 1; i <= counter; i++)
	{
		txttemp1 = document.getElementById("DataDouble"+i+"");
		txttemp1.value = stringfix(txttemp1.value);
		if(txttemp1.value == "")
		{
			if(bool == true)
			{
				loge.innerHTML = "Please Make Sure all your Fields have a Name.";
				txttemp1.focus();
			}
			bool = false;
			txttemp1.style.borderColor = "#7e3d3f";
		}
		else
		{
			if(stringcheck(txttemp1.value) == false)
			{
				loge.innerHTML = error;
				txttemp1.value = stringfix(txttemp1.value);
				bool = false;
			}
			else
				if(txttemp1.style.borderColor != "#7e3d3f")
						txttemp1.style.backgroundColor = "#476A34";	
			
			txttemp1.value = stringfix(txttemp1.value);
		}
		for(var j = 1; j <= counter; j++)
		{
			if(i == j && j <= counter)
				j++;
			if(j > counter)
				break;
			txttemp2 = document.getElementById("DataDouble"+j+"");
			if(txttemp1.value == txttemp2.value && txttemp2.value != "")
			{
				txttemp2.style.borderColor = "#7e3d3f";
				if(bool == true)
				{
					loge.innerHTML = "Please Have Separate Names for all of your Fields.<br/>Please Look at the fields outlined in Red.";
					txttemp2.focus();
				}
				bool = false;
			}
			else
				if(txttemp2.style.borderColor != "#7e3d3f")
					txttemp2.style.backgroundColor = "#476A34";
		}
	}
	//checks currency fields agaisnt currency fields
	var counter = document.getElementById("curra").value;
	for(var i = 1; i <= counter; i++)
	{
		txttemp1 = document.getElementById("DataCurr"+i+"");
		txttemp1.value = stringfix(txttemp1.value);
		if(txttemp1.value == "")
		{
			if(bool == true)
			{
				loge.innerHTML = "Please Make Sure all your Fields have a Name.";
				txttemp1.focus();
			}
			bool = false;
			txttemp1.style.borderColor = "#7e3d3f";
		}
		else
		{
			if(stringcheck(txttemp1.value) == false)
			{
				loge.innerHTML = error;
				txttemp1.value = stringfix(txttemp1.value);
				bool = false;
			}
			else
				if(txttemp1.style.borderColor != "#7e3d3f")
						txttemp1.style.backgroundColor = "#476A34";	
			
			txttemp1.value = stringfix(txttemp1.value);
		}
		for(var j = 1; j <= counter; j++)
		{
			if(i == j && j <= counter)
				j++;	
			if(j > counter)
				break;	
			txttemp2 = document.getElementById("DataCurr"+j+"");
			if(txttemp1.value == txttemp2.value && txttemp2.value != "")
			{
				txttemp2.style.borderColor = "#7e3d3f";
				if(bool == true)
				{
					loge.innerHTML = "Please Have Separate Names for all of your Fields.<br/>Please Look at the fields outlined in Red.";
					txttemp2.focus();
				}
				bool = false;
			}
			else
				if(txttemp2.style.borderColor != "#7e3d3f")
					txttemp2.style.backgroundColor = "#476A34";
		}
	}
	//checks date fields agaisnt date fields
	var counter = document.getElementById("datea").value;
	for(var i = 1; i <= counter; i++)
	{
		txttemp1 = document.getElementById("DataDate"+i+"");
		txttemp1.value = stringfix(txttemp1.value);
		if(txttemp1.value == "")
		{
			if(bool == true)
			{
				loge.innerHTML = "Please Make Sure all your Fields have a Name.";
				txttemp1.focus();
			}
			bool = false;
			txttemp1.style.borderColor = "#7e3d3f";
		}
		else
		{
			if(re.test(txttemp1.value))
			{
				loge.innerHTML = "Column names can't start with a number(Please See the Administrator Manual For More Information)";
				bool = false;
			}
			else if(stringcheck(txttemp1.value) == false)
			{
				loge.innerHTML = error;
				bool = false;
			}
			else
			{
				if(txttemp1.style.borderColor != "#7e3d3f")
						txttemp1.style.backgroundColor = "#476A34";	
			}
		}
		for(var j = 1; j <= counter; j++)
		{
			if(i == j && j <= counter)
				j++;
			if(j > counter)
				break;	
			txttemp2 = document.getElementById("DataDate"+j+"");
			if(txttemp1.value == txttemp2.value && txttemp2.value != "")
			{
				txttemp2.style.borderColor = "#7e3d3f";
				if(bool == true)
				{
					loge.innerHTML = "Please Have Separate Names for all of your Fields.<br/>Please Look at the fields outlined in Red.";
					txttemp2.focus();
				}
				bool = false;
			}
			else
				if(txttemp2.style.borderColor != "#7e3d3f")
					txttemp2.style.backgroundColor = "#476A34";
		}
	}
	//checks check boxes agaisnt check boxes
	var counter = document.getElementById("checka").value;
	for(var i = 1; i <= counter; i++)
	{
		txttemp1 = document.getElementById("DataCheck"+i+"");
		txttemp1.value = stringfix(txttemp1.value);
		if(txttemp1.value == "")
		{
			if(bool == true)
			{
				loge.innerHTML = "Please Make Sure all your Fields have a Name.";
				txttemp1.focus();
			}
			bool = false;
			txttemp1.style.borderColor = "#7e3d3f";
		}
		else
		{
			if(stringcheck(txttemp1.value) == false)
			{
				loge.innerHTML = error;
				txttemp1.value = stringfix(txttemp1.value);
				bool = false;
			}
			else
				if(txttemp1.style.borderColor != "#7e3d3f")
						txttemp1.style.backgroundColor = "#476A34";	
			
			txttemp1.value = stringfix(txttemp1.value);
		}
		for(var j = 1; j <= counter; j++)
		{
			if(i == j && j <= counter)
				j++;	
			if(j > counter)
				break;
			txttemp2 = document.getElementById("DataCheck"+j+"");
			if(txttemp1.value == txttemp2.value && txttemp2.value != "")
			{
				txttemp2.style.borderColor = "#7e3d3f";
				if(bool == true)
				{
					loge.innerHTML = "Please Have Separate Names for all of your Fields.<br/>Please Look at the fields outlined in Red.";
					txttemp2.focus();
				}
				bool = false;
			}
			else
				if(txttemp2.style.borderColor != "#7e3d3f")
				txttemp2.style.backgroundColor = "#476A34";
		}
	}
	//checks drop downs against drop downs
	var counter = document.getElementById("dropa").value;
	for(var i = 1; i <= counter; i++)
	{
		txttemp1 = document.getElementById("SDropName"+i+"");
		txttemp1.value = stringfix(txttemp1.value);
		if(txttemp1.value == "")
		{
			if(bool == true)
			{
				loge.innerHTML = "Please Make Sure all your Fields have a Name.";
				txttemp1.focus();
			}
			bool = false;
			txttemp1.style.borderColor = "#7e3d3f";
		}
		else
		{
			if(stringcheck(txttemp1.value) == false)
			{
				loge.innerHTML = error;
				txttemp1.value = stringfix(txttemp1.value);
				bool = false;
			}
			else
				if(txttemp1.style.borderColor != "#7e3d3f")
						txttemp1.style.backgroundColor = "#476A34";	
			
			txttemp1.value = stringfix(txttemp1.value);
		}
		for(var j = 1; j <= counter; j++)
		{
			if(i == j && j <= counter)
				j++;
			if(j > counter)
				break;	
			txttemp2 = document.getElementById("SDropName"+j+"");
			if(txttemp1.value == txttemp2.value && txttemp2.value != "")
			{
				txttemp2.style.borderColor = "#7e3d3f";
				if(bool == true)
				{
					loge.innerHTML = "Please Have Separate Names for all of your Fields.<br/>Please Look at the fields outlined in Red.";
					txttemp2.focus();
				}
				bool = false;
			}
			else
				if(txttemp2.style.borderColor != "#7e3d3f")
					txttemp2.style.backgroundColor = "#476A34";
		}
	}	
	//checksHyperlink fields against Hyperlink fields
	var counter = document.getElementById("hypera").value;
	for(var i = 1; i <= counter; i++)
	{
		txttemp1 = document.getElementById("DataHyper"+i+"");
		txttemp1.value = stringfix(txttemp1.value);
		if(txttemp1.value == "")
		{
			if(bool == true)
			{
				loge.innerHTML = "Please Make Sure all your Fields have a Name.";
				txttemp1.focus();
			}
			bool = false;
			txttemp1.style.borderColor = "#7e3d3f";
		}
		else
		{
			if(stringcheck(txttemp1.value) == false)
			{
				loge.innerHTML = error;
				txttemp1.value = stringfix(txttemp1.value);
				bool = false;
			}
			else
				if(txttemp1.style.borderColor != "#7e3d3f")
						txttemp1.style.backgroundColor = "#476A34";	
			
			txttemp1.value = stringfix(txttemp1.value);
		}
		for(var j = 1; j <= counter; j++)
		{
			if(i == j && j <= counter)
				j++;
			if(j > counter)
				break;
			txttemp2 = document.getElementById("DataHyper"+j+"");
			if(txttemp1.value == txttemp2.value && txttemp2.value != "")
			{
				txttemp2.style.borderColor = "#7e3d3f";
				if(bool == true)
				{
					loge.innerHTML = "Please Have Separate Names for all of your Fields.<br/>Please Look at the fields outlined in Red.";
					txttemp2.focus();
				}
				bool = false;
			}
			else
				if(txttemp2.style.borderColor != "#7e3d3f")
					txttemp2.style.backgroundColor = "#476A34";
		}
	}
	//checks Autofill fields against Autofill fields
	var counter = document.getElementById("autofa").value;
	for(var i = 1; i <= counter; i++)
	{
		txttemp1 = document.getElementById("AutoName"+i+"");
		txttemp1.value = stringfix(txttemp1.value);
		if(txttemp1.value == "")
		{
			if(bool == true)
			{
				loge.innerHTML = "Please Make Sure all your Fields have a Name.";
				txttemp1.focus();
			}
			bool = false;
			txttemp1.style.borderColor = "#7e3d3f";
		}
		else
		{
			if(stringcheck(txttemp1.value) == false)
			{
				loge.innerHTML = error;
				txttemp1.value = stringfix(txttemp1.value);
				bool = false;
			}
			else
				if(txttemp1.style.borderColor != "#7e3d3f")
						txttemp1.style.backgroundColor = "#476A34";	
			
			txttemp1.value = stringfix(txttemp1.value);
		}
		for(var j = 1; j <= counter; j++)
		{
			if(i == j && j <= counter)
				j++;
			if(j > counter)
				break;
			txttemp2 = document.getElementById("AutoName"+j+"");
			if(txttemp1.value == txttemp2.value && txttemp2.value != "")
			{
				txttemp2.style.borderColor = "#7e3d3f";
				if(bool == true)
				{
					loge.innerHTML = "Please Have Separate Names for all of your Fields.<br/>Please Look at the fields outlined in Red.";
					txttemp2.focus();
				}
				bool = false;
			}
			else
				if(txttemp2.style.borderColor != "#7e3d3f")
					txttemp2.style.backgroundColor = "#476A34";
		}
	}
	return bool;
	//
	//End of try start of catch
	//
	}
	catch(e)
	{
		var e1 = "checkadmin.js had an error in testallcases4CNT() function";
		ErrorToFile(e, e1);
		window.location.reload();
	}
}
//change column names
function CHeckCN()
{
	//
	//Start of Try Function I wont tab here
	//
	try{
	var loge = document.getElementById("loge");
	var getTN = document.getElementById("OTableName");
	var TN = getTN.options[getTN.selectedIndex].text;
	TN = stringfix(TN)
	var myDB = new ACCESSdb("Abcon.mdb", {showErrors:true});
	var counter = document.getElementById("texta").value;
	for(var i = 1; i <= counter;i++)
	{
		var oldcolname = document.getElementById("DataTextO"+i+"").value
		var newcolname = document.getElementById("DataText"+i+"").value
		newcolname = stringfix(newcolname);
		if(newcolname != "")
		{
			if(oldcolname != newcolname)
			{
				var SQL = "ALTER TABLE ["+TN+"] ADD COLUMN ["+newcolname+"] varchar(255)";
				myDB.query(SQL);
				SQL = "UPDATE ["+TN+"] SET ["+newcolname+"] = ["+oldcolname+"]";
				myDB.query(SQL);
				SQL = "ALTER TABLE ["+TN+"] DROP COLUMN ["+oldcolname+"]";
				myDB.query(SQL);
				var SQL = "Update Lookuptable Set GetFieldName='"+(newcolname)+"' Where GetFieldName='"+(oldcolname)+"'";
				myDB.query(SQL);
				var SQL = "Update Lookuptable Set RecFieldName='"+(newcolname)+"' Where RecFieldName='"+(oldcolname)+"'";
				myDB.query(SQL);
			}
		}
	}
	var counter = document.getElementById("memoa").value;
	for(var i = 1; i <= counter;i++)
	{
		var oldcolname = document.getElementById("DataMemoO"+i+"").value
		var newcolname = document.getElementById("DataMemo"+i+"").value
		newcolname = stringfix(newcolname);
		if(newcolname != "")
		{
			if(oldcolname != newcolname)
			{
				var SQL = "ALTER TABLE ["+TN+"] ADD COLUMN ["+newcolname+"] MEMO";
				myDB.query(SQL);
				SQL = "UPDATE ["+TN+"] SET ["+newcolname+"] = ["+oldcolname+"]";
				myDB.query(SQL);
				SQL = "ALTER TABLE ["+TN+"] DROP COLUMN ["+oldcolname+"]";
				myDB.query(SQL);
			}
		}
	}
	var counter = document.getElementById("inta").value;
	for(var i = 1; i <= counter;i++)
	{
		var oldcolname = document.getElementById("DataIntO"+i+"").value
		var newcolname = document.getElementById("DataInt"+i+"").value
		newcolname = stringfix(newcolname);
		if(newcolname != "")
		{
			if(oldcolname != newcolname)
			{
				var SQL = "ALTER TABLE ["+TN+"] ADD COLUMN ["+newcolname+"] Long";
				myDB.query(SQL);
				SQL = "UPDATE ["+TN+"] SET ["+newcolname+"] = ["+oldcolname+"]";
				myDB.query(SQL);
				SQL = "ALTER TABLE ["+TN+"] DROP COLUMN ["+oldcolname+"]";
				myDB.query(SQL);
			}
		}
	}
	var counter = document.getElementById("doublea").value;
	for(var i = 1; i <= counter;i++)
	{
		var oldcolname = document.getElementById("DataDoubleO"+i+"").value
		var newcolname = document.getElementById("DataDouble"+i+"").value
		newcolname = stringfix(newcolname);
		if(newcolname != "")
		{
			if(oldcolname != newcolname)
			{
				var SQL = "ALTER TABLE ["+TN+"] ADD COLUMN ["+newcolname+"] Double";
				myDB.query(SQL);
				SQL = "UPDATE ["+TN+"] SET ["+newcolname+"] = ["+oldcolname+"]";
				myDB.query(SQL);
				SQL = "ALTER TABLE ["+TN+"] DROP COLUMN ["+oldcolname+"]";
				myDB.query(SQL);
			}
		}
	}
	var counter = document.getElementById("curra").value;
	for(var i = 1; i <= counter;i++)
	{
		var oldcolname = document.getElementById("DataCurrO"+i+"").value
		var newcolname = document.getElementById("DataCurr"+i+"").value
		newcolname = stringfix(newcolname);
		if(newcolname != "")
		{
			if(oldcolname != newcolname)
			{
				var SQL = "ALTER TABLE ["+TN+"] ADD COLUMN ["+newcolname+"] Currency";
				myDB.query(SQL);
				SQL = "UPDATE ["+TN+"] SET ["+newcolname+"] = ["+oldcolname+"]";
				myDB.query(SQL);
				SQL = "ALTER TABLE ["+TN+"] DROP COLUMN ["+oldcolname+"]";
				myDB.query(SQL);
			}
		}
	}
	var counter = document.getElementById("datea").value;
	for(var i = 1; i <= counter;i++)
	{
		var oldcolname = document.getElementById("DataDateO"+i+"").value
		var newcolname = document.getElementById("DataDate"+i+"").value
		newcolname = stringfix(newcolname);
		if(newcolname != "")
		{
			if(oldcolname != newcolname)
			{
				var SQL = "ALTER TABLE ["+TN+"] ADD COLUMN ["+newcolname+"] DATETIME";
				myDB.query(SQL);
				SQL = "UPDATE ["+TN+"] SET ["+newcolname+"] = ["+oldcolname+"]";
				myDB.query(SQL);
				SQL = "ALTER TABLE ["+TN+"] DROP COLUMN ["+oldcolname+"]";
				myDB.query(SQL);
			}
		}
	}
	var counter = document.getElementById("checka").value;
	for(var i = 1; i <= counter;i++)
	{
		var oldcolname = document.getElementById("DataCheckO"+i+"").value
		var newcolname = document.getElementById("DataCheck"+i+"").value
		newcolname = stringfix(newcolname);
		if(newcolname != "")
		{
			if(oldcolname != newcolname)
			{
				var SQL = "ALTER TABLE ["+TN+"] ADD COLUMN ["+newcolname+"] YESNO";
				myDB.query(SQL);
				SQL = "UPDATE ["+TN+"] SET ["+newcolname+"] = ["+oldcolname+"]";
				myDB.query(SQL);
				SQL = "ALTER TABLE ["+TN+"] DROP COLUMN ["+oldcolname+"]";
				myDB.query(SQL);
			}
		}
	}
	var counter = document.getElementById("dropa").value;
	for(var i = 1; i <= counter;i++)
	{
		var oldcolname = document.getElementById("SDropNameO"+i+"").value
		var newcolname = document.getElementById("SDropName"+i+"").value
		newcolname = stringfix(newcolname);
		if(newcolname != "")
		{
			if(oldcolname != newcolname)
			{
				var SQL = "ALTER TABLE ["+TN+"] ADD COLUMN ["+newcolname+"] varchar(255)";
				myDB.query(SQL);
				SQL = "UPDATE ["+TN+"] SET ["+newcolname+"] = ["+oldcolname+"]";
				myDB.query(SQL);
				SQL = "ALTER TABLE ["+TN+"] DROP COLUMN ["+oldcolname+"]";
				myDB.query(SQL);
				var SQL = "Update Lookuptable Set GetFieldName='"+(newcolname)+"' Where GetFieldName='"+(oldcolname)+"'";
				myDB.query(SQL);
				var SQL = "Update Lookuptable Set RecFieldName='"+(newcolname)+"' Where RecFieldName='"+(oldcolname)+"'";
				myDB.query(SQL);
			}
		}
	}
	var counter = document.getElementById("hypera").value;
	for(var i = 1; i <= counter;i++)
	{
		var oldcolname = document.getElementById("SDropNameO"+i+"").value
		var newcolname = document.getElementById("SDropName"+i+"").value
		newcolname = stringfix(newcolname);
		if(newcolname != "")
		{
			if(oldcolname != newcolname)
			{
				var SQL = "ALTER TABLE ["+TN+"] ADD COLUMN ["+newcolname+"] varchar(255)";
				myDB.query(SQL);
				SQL = "UPDATE ["+TN+"] SET ["+newcolname+"] = ["+oldcolname+"]";
				myDB.query(SQL);
				SQL = "ALTER TABLE ["+TN+"] DROP COLUMN ["+oldcolname+"]";
				myDB.query(SQL);
				var SQL = "Update HyperlinkCheck Set FieldName='"+(newcolname)+"' Where FieldName='"+(oldcolname)+"'";
				myDB.query(SQL);
			}
		}
	}
	var counter = document.getElementById("autofa").value;
	for(var i = 1; i <= counter;i++)
	{
		var oldcolname = document.getElementById("AutoNameO"+i+"").value
		var newcolname = document.getElementById("AutoName"+i+"").value
		newcolname = stringfix(newcolname);
		if(newcolname != "")
		{
			if(oldcolname != newcolname)
			{
				var SQL = "ALTER TABLE ["+TN+"] ADD COLUMN ["+newcolname+"] varchar(255)";
				myDB.query(SQL);
				SQL = "UPDATE ["+TN+"] SET ["+newcolname+"] = ["+oldcolname+"]";
				myDB.query(SQL);
				SQL = "ALTER TABLE ["+TN+"] DROP COLUMN ["+oldcolname+"]";
				myDB.query(SQL);
				var SQL = "Update Lookuptable Set GetFieldName='"+(newcolname)+"' Where GetFieldName='"+(oldcolname)+"'";
				myDB.query(SQL);
				var SQL = "Update Lookuptable Set RecFieldName='"+(newcolname)+"' Where RecFieldName='"+(oldcolname)+"'";
				myDB.query(SQL);
			}
		}
	}
		window.location.reload();
	}
		//
		//End of Try Start of Catch
		//
	catch(e)
	{
		var e1 = "checkadmin.js had an error in CHeckCN() function";
		ErrorToFile(e, e1);
		window.location.reload();
	}
}
//adds attachments like forms pdf etc to the main page for download

function addAttachment(){
	try{
		var attName = document.getElementById("atttitle").value;
		var temp = document.getElementById("Attachment4MP");
		var aray = temp.innerHTML;
		var myObject = new ActiveXObject("Scripting.FileSystemObject");
		var origurl = temp.value;
		if(temp.value == "")
		{
			origurl = aray;
		}
		if( origurl != "")
		{
			var filename = origurl.substring(origurl.lastIndexOf("\\")+1);
			var finaldir = "\DatabaseFiles\\"+filename+"";
			myObject.CopyFile (origurl, finaldir);
		}
		var myDB = new ACCESSdb("Abcon.mdb", {showErrors:false});
		var SQL = "Insert into Attachments(FileName, Att) VALUES('"+attName+"', '"+filename+"')";
		myDB.query(SQL);
	}
	catch(e)
	{
		var e1 = "checkadmin.js had an error in addAttachment() function";
		ErrorToFile(e, e1);
		window.location.reload();
	}
}