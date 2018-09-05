//this portion of javascript code fills in  the page adminpanel.hta with the form data to fill the popup called pop4form
//THis portion of code fills a div with the form to add a process
//no need to use try catch errors here an error should not occur
function fillAddProcess()
{
	var getp4f = document.getElementById("fillp4f");
	getp4f.innerHTML = "<div id='loge' class='errormessage'></div>"+
	"<div class='twelve columns$'>"+
	"<label for='demo-name' class='Title4Form'>Form to Add Process</label>"+
	"<div id='loge' class='errormessage'></div>"+
	"<form method='post' onsubmit='return anprocess()'>"+
	"<div class='u-full-width'>"+
		"<label for='demo-name'>Title</label>"+
		"<input type='text' name='demo-name' id='proctitle' value='' />"+
		"<label for='demo-message'>Information</label>"+
		"<textarea name='demo-message' id='procinfo' rows='6'></textarea>"+
		"<ul class='actions'>"+
			"<li class='six columns'><input type='submit' value='Add Process' class='special' /></li>"+
			"<li class='six columns'><input type='reset' value='Reset Form' /></li>"+
		"</ul></div></form><a id='close' href='#' class='close'>&nbsp;&nbsp;&nbsp;</a>";
};
//this returns the dropdown menus for Edit , Delete, Hide Process aswell as create a ta
function returnoptions4dd(ParentID)
{
	try{
		var adoConn = new ActiveXObject("ADODB.Connection");
		var adoRS = new ActiveXObject("ADODB.Recordset");
		adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
		var temp = false;
		if(ParentID == null)
			adoRS.Open("Select * From HomePage", adoConn);
		else
			adoRS.Open("Select * From HomePage Where  ID <> ["+ParentID+"]", adoConn);
		if(!adoRS.bof)
		{
			//fills out the options in the selection box only if there are elements to add 
			//otherwise asks if you want to add a process
	//this one doesnt use accessdb.js I needed more fucntionality so I did it myself
			adoRS.MoveFirst;
			temp = "<option value=0></option>";
			while(!adoRS.eof)
			{
				temp += "<option id="+adoRS.fields(0)+">"+(adoRS.fields(1).value)+"</options>";
				adoRS.movenext;
			}
			adoRS.Close();
			adoConn.Close();
		}
		return temp;
	}
	catch(e)
	{
		var e1 = "fillformdata.js had an error in the returnoptions4dd function";
		ErrorToFile(e, e1);
		location.reload();
	}
}
//returns a drop down off all avaialble tables
function returnoptions4CT()
{
	try{
		var adoConn = new ActiveXObject("ADODB.Connection");
		var adoRS = new ActiveXObject("ADODB.Recordset");
		adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
		var temp = false;
		adoRS.Open("Select * From TableList", adoConn);
		if(!adoRS.bof)
		{
			//fills out the options in the selection box only if there are elements to add 
			//otherwise asks if you want to add a process
			//this one doesn't use accessdb.js I needed more functionality so I did it myself
			adoRS.MoveFirst;
			temp = "<option value=0></option>";
			while(!adoRS.eof)
			{
				temp += "<option id="+adoRS.fields(0)+" value="+adoRS.fields(2)+">"+adoRS.fields(1).value+"</options>";
				adoRS.movenext;
			}
			adoRS.Close();
			adoConn.Close();
		}
		return temp;
	}
	catch(e)
	{
		var e1 = "fillformdata.js had an error in the returnoptions4CT function";
		ErrorToFile(e, e1);
		location.reload();
	}
}

//fills a form to edit a process if there are no processes it asks if you want to make one
function fillEditProcess()
{
	var temp = returnoptions4dd();
	if(temp != false)
	{
		var getp4f = document.getElementById("fillp4f");
		getp4f.innerHTML = "<div id='loge' class='errormessage'></div>"+
		"<label for='demo-name' class='Title4Form'>What Process Do You Want to Edit</label>"+
		"<div id='loge' class='errormessage'></div><label for='demo-name'>Plese Select A Title</label>"+
		"<div class='select-style'><select id='ptsel' onchange='return callfromctitle()'>"+temp+"</select></div><div id='fillp4f2'></div><a id='close' href='#' class='close'>&nbsp;&nbsp;&nbsp;</a>";
	}
	else{
		notavailP();
	}
};
// just populates the rest of the form to edit the process
function callfromctitle()
{
	var ptsel = document.getElementById("ptsel");
	var getp4f = document.getElementById("fillp4f2");
	var test = ptsel.options[ptsel.selectedIndex].value;
	getp4f.innerHTML ="<form method='post' onsubmit='return eoprocess()'>"+
	"<label for='demo-name'>Title</label>"+
	"<input type='text' name='"+ptsel.options[ptsel.selectedIndex].id+"' id='eproctitle' value='"+ptsel.options[ptsel.selectedIndex].text+"' />"+
	"<label for='demo-message'>Information</label>"+
	"<textarea name='demo-message' id='eprocinfo' rows='6'></textarea>"+
	"<ul class='actions'>"+
		"<li class='six columns'><input type='submit' value='Edit Process' class='special' /></li>"+
		"<li class='six columns'><input type='reset' value='Reset Form' /></li></ul></form>";
};
function notavailP()
{
		var getp4f = document.getElementById("fillp4f");
		getp4f.innerHTML = "<div id='loge' class='errormessage'></div>";
		var loge = document.getElementById("loge");
		loge.innerHTML = "The Table is Empty";
		loge.style.color = "Blue";
		loge.style.backgroundColor = "#fee0c6";
		loge.style.borderColor = "#7e3d3f";
		getp4f.innerHTML +=	"<label for='demo-name' class='Title4Form'>Do you want to Add a Process</label>"+
		"<div id='loge' class='errormessage'></div>"+
		"<ul class='actions'>"+
			"<li class='six columns'><input type='submit' onclick='return fillAddProcess()' value='Yes' class='special' /></li>"+
			"<li class='six columns'><input type='reset' onclick='return backtopanel()' value='No' /></li></ul><a id='close' href='#' class='close'>&nbsp;&nbsp;&nbsp;</a>";
}
//used to delete a process pretty much the exact same as the first part of the edit process
function DelProcess()
{
	var temp = returnoptions4dd();
	if(temp != false)
	{
		var getp4f = document.getElementById("fillp4f");
		getp4f.innerHTML = "<div id='loge' class='errormessage'></div>"+
		"<label for='demo-name' class='Title4Form'>What Process Do You Want to Delete</label>"+
		"<div id='loge' class='errormessage'></div><label for='demo-name'>Plese Select A Title</label>"+
		"<div class='select-style'><select id='ptsel' onchange='return DeleteProcessC()'>"+temp+"</select></div><div id='fillp4f2'></div><a id='close' href='#' class='close'>&nbsp;&nbsp;&nbsp;</a>";
	}
	else{
		notavailP();
	}
};
//populates the continued form to delete the chosen process
function DeleteProcessC()
{
	var ptsel = document.getElementById("ptsel");
	var getp4f = document.getElementById("fillp4f2");
	getp4f.innerHTML ="<form method='post' action='#' onsubmit='return doprocess()'>"+
	"<div class='twelve columns$'>"+
		"<label for='demo-name'>Are you sure you want to Delete Process</label>"+
		"<input type='text' name='"+ptsel.options[ptsel.selectedIndex].id+"' id='eproctitle' value='"+ptsel.options[ptsel.selectedIndex].text+"' readonly/>"+
	"</div>"+
	"<div class='twelve columns$'>"+
		"<ul class='actions'>"+
			"<li class='six columns'><input type='submit' value='Confirm Deletion' class='special' /></li>"+
			"<li class='six columns'><input type='reset' onclick='return DelProcess()' value='Back' /></li>"+
		"</ul></div></div></form>";
};
//Hides a process if it no longer is used still visible in this menu though also very similar to the edit command
function HideProcess()
{
	var temp = returnoptions4dd();
	if(temp != false)
	{
		var getp4f = document.getElementById("fillp4f");
		getp4f.innerHTML = "<div id='loge' class='errormessage'></div>"+
		"<label for='demo-name' class='Title4Form'>What Process Do You Want to Toggle Visibility for</label>"+
		"<div id='loge' class='errormessage'></div><label for='demo-name'>Plese Select A Title</label>"+
		"<div class='select-style'><select id='ptsel' onchange='return HideProcessC()'>"+temp+"</select></div><div id='fillp4f2'></div><a id='close' href='#' class='close'>&nbsp;&nbsp;&nbsp;</a>";
	}
	else{
		notavailP();
	}
};
//populates the rest of the toggle visibility process
function HideProcessC()
{
	var ptsel = document.getElementById("ptsel");
	var getp4f = document.getElementById("fillp4f2");
	getp4f.innerHTML ="<form method='post' action='#' onsubmit='return hoprocess()'>"+
	"<label for='demo-name'>Are you sure you want to Toggle Visibility</label>"+
	"<input type='text' name='"+ptsel.options[ptsel.selectedIndex].id+"' id='eproctitle' value='"+ptsel.options[ptsel.selectedIndex].text+"' readonly/>"+
	"<ul class='actions'>"+
		"<li class='six columns'><input type='submit' value='Confirm Hide' class='special' /></li>"+
		"<li class='six columns'><input type='reset' onclick='return HideProcess()' value='Back' /></li>"+
	"</ul></form>";
};
// Starts the Process to Create A Table Start is very similar to edit process since we are/
//creating the table for a process
function TableCreate()
{
	var temp = returnoptions4dd();
	if(temp != false)
	{
		var getp4f = document.getElementById("fillp4f");
		getp4f.innerHTML = "<div id='loge' class='errormessage'></div>"+
		"<label for='demo-name' class='Title4Form'>What Process Do You Want to Create the Table for</label>"+
		"<div id='loge' class='errormessage'></div><label for='demo-name'>Plese Select A Process</label>"+
		"<div class='select-style'><select id='ptsel' onchange='return TableCreateC()'>"+temp+"</select></div><div id='fillp4f2'></div><a id='close' href='#' class='close'>&nbsp;&nbsp;&nbsp;</a>";
	}
	else{
		notavailP();
	}
};
//this iws where we 
function TableCreateC()
{
	var ptsel = document.getElementById("ptsel");
	var getp4f = document.getElementById("fillp4f2");
	getp4f.innerHTML ="<form method='post'>"+
		"<label for='demo-name'>What do you want to name your table</label>"+
		"<input name='"+ptsel.options[ptsel.selectedIndex].id+"' id='tableName' value='' type='text'/>"+
	"</div><div id='fillp4f3'><ul class='actions'>"+
		"<li class='six columns'><input type='submit' onclick='return EditCreateC()' value='Confirm Table Name' class='special' /></li>"+
		"<li class='six columns'><input type='reset' onclick='return TableCreateC()' value='Reset' /></li>"+
	"</ul></div></div>"+
	"</form>";
};
//checks whether it is an autofill or a drop-down then fills up the drop down or autofill posibilities
function textboxcheck(TN, FN)
{
	try{
		var adoConn = new ActiveXObject("ADODB.Connection");
		var adoRS = new ActiveXObject("ADODB.Recordset");
		adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
		adoRS.Open("Select * From Lookuptable Where RecTableName='["+TN+"]' AND RecFieldName='"+FN+"'", adoConn);
		if(!adoRS.bof)
		{
			adoRS.MoveFirst;
			if(!adoRS.eof)
			{
				if(adoRS.fields("Type").value == 'D')
				{
					var getCD = AddDropCheck();
					var DDTemp = document.getElementById("SelTDrop"+getCD+"");
					DDTemp.options[0].text = adoRS.fields("GetTableName").value;
					dropdownmeu(getCD);
					DDTemp = document.getElementById("SelFDrop"+getCD+"");
					DDTemp.options[0].text = adoRS.fields("GetFieldName").value;
					document.getElementById("SDropName"+getCD+"").value = FN;
					document.getElementById("SDropNameO"+getCD+"").value = FN;
				}
				else if(adoRS.fields("Type").value == 'A')
				{
					var getCA = AddAFillCheck();
					var DDTemp = document.getElementById("SelATDrop"+getCA+"");
					DDTemp.options[0].text = adoRS.fields("GetTableName").value;
					dropdownmeu(getCA,1);
					DDTemp = document.getElementById("SelAFDrop"+getCA+"");
					DDTemp.options[0].text = adoRS.fields("GetFieldName").value;
					document.getElementById("AutoName"+getCA+"").value = FN;
					document.getElementById("AutoNameO"+getCA+"").value = FN;
				}
			}
		}
		else{
			var value = AddDataText();
			document.getElementById("DataTextO"+value+"").value = FN;
			document.getElementById("DataText"+value+"").value = FN;
		}
		
		adoRS.Close();
		adoConn.Close();
	}
	catch(e)
	{
		var e1 = "fillformdata.js had an error in the  textboxcheck(TN, FN) function";
		ErrorToFile(e, e1);
		location.reload();
	}
}
//checks if column should be memoboxcheck
function memoboxcheck(TN,FN)
{
	try{
		var adoConn = new ActiveXObject("ADODB.Connection");
		var adoRS = new ActiveXObject("ADODB.Recordset");
		adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
		adoRS.Open("Select * From HyperlinkCheck Where TableName='"+TN+"' AND FieldName='"+FN+"'", adoConn);
		if(!adoRS.bof)
		{
			var value = AddHyperCheck();
			if(adoRS.fields("TableName").value == TN)
			{
				document.getElementById("DataHyper"+value+"").value = FN;
				document.getElementById("DataHyperO"+value+"").value = FN;
			}
		}
		else{
			var value = AddDataMemo();
			document.getElementById("DataMemoO"+value+"").value = FN;
			document.getElementById("DataMemo"+value+"").value = FN;
		}
		
		adoRS.Close();
		adoConn.Close();
	}
	catch(e)
	{
		var e1 = "fillformdata.js had an error in the memoboxcheck(TN, FN) function";
		ErrorToFile(e, e1);
		location.reload();
	}
}
//loads previous data into the change column name option
function loadoldtableinfo()
{
	try{
		var SelectOTN = document.getElementById("OTableName");
		var TableName = stringfix(SelectOTN.options[SelectOTN.selectedIndex].text);
		var adoConn = new ActiveXObject("ADODB.Connection");
		var adoRS = new ActiveXObject("ADODB.Recordset");
		adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
		adoRS.Open("Select * From ["+TableName+"]", adoConn);
		for(var i = 1; i < adoRS.Fields.Count; i++)
		{
			var temp = adoRS.fields(i).type;
			switch(temp)
			{
				case 202:	//text autofill and dropdown
							textboxcheck(TableName, adoRS.fields(i).name);
							break;
				case 203:	//memo and hyperlink
							memoboxcheck(TableName, adoRS.fields(i).name);
							break;
				case 3:		var value = AddDataInt();
							document.getElementById("DataInt"+value+"").value = adoRS.fields(i).name;
							document.getElementById("DataIntO"+value+"").value = adoRS.fields(i).name;
							break;
				case 5:		var value = AddDataDouble();
							document.getElementById("DataDouble"+value+"").value = adoRS.fields(i).name;
							document.getElementById("DataDoubleO"+value+"").value = adoRS.fields(i).name;
							break;
				case 6:		var value = AddDataCurr();
							document.getElementById("DataCurr"+value+"").value = adoRS.fields(i).name;
							document.getElementById("DataCurrO"+value+"").value = adoRS.fields(i).name;
							break;
				case 7:		var value = AddDataDate();
							document.getElementById("DataDate"+value+"").value = adoRS.fields(i).name;
							document.getElementById("DataDateO"+value+"").value = adoRS.fields(i).name;
							break;
				case 11:	var value = AddDataCheck();
							document.getElementById("DataCheck"+value+"").value = adoRS.fields(i).name;
							document.getElementById("DataCheckO"+value+"").value = adoRS.fields(i).name;
							break;
			}
		}
		adoRS.Close();
		adoConn.Close();
	}
	catch(e)
	{
		var e1 = "fillformdata.js had an error in the  loadoldtableinfo() function";
		ErrorToFile(e, e1);
		location.reload();
	}
}
//loads load old table info function after finishing populating the basics of the table
function ChangeColumnNames()
{
	var getp4f = document.getElementById("fillp4f3");
	getp4f.innerHTML ="<div id='addbox'></div><input type='hidden' id='totccnt'  value=0>"+
	"<input type='hidden' id='texta' value=0 name='test'>"+
	"<input type='hidden' id='memoa' value=0 name='test'>"+
	"<input type='hidden' id='inta' value=0 name='test'>"+
	"<input type='hidden' id='doublea' value=0 name='test'>"+
	"<input type='hidden' id='curra' value=0 name='test'>"+
	"<input type='hidden' id='datea' value=0 name='test'>"+
	"<input type='hidden' id='checka' value=0 name='test'>"+
	"<input type='hidden' id='dropa' value=0 name='test'>"+
	"<input type='hidden' id='hypera'  value=0 name='test'>"+
	"<input type='hidden' id='autofa'  value=0 name='test'>"+
	"<div class='twelve columns$'>"+
		"<ul class='actions'>"+
			"<li class='six columns'><input type='submit' onclick='return CHeckCN()' value='Add Table' class='special' /></li>"+
			"<li class='six columns'><input type='reset' onclick='return ChangeColumnNames()' value='Reset Form' /></li>"+
		"</ul>"+
	"</div></div>"+
	"</form></div>";
	loadoldtableinfo();
}
//this is where the adding of the table begins we can add various columns and datatypes in this row.
// all the fill out boxes are automatically populated by clicking the radio buttons
//might add a hover card to explain more details of each column might need to add more datatypes but for now this looks good
function EditCreateC(alter)
{
	var getp4f = document.getElementById("fillp4f3");
	getp4f.innerHTML ="<div id='addbox'></div>"+
	"<div class='twelve columns'><label for='demo-category'>What Data Type do you want to add to the table?</label>"+
	"<div class='two columns'>"+
	"<input type='hidden' id='totccnt'  value=0>"+
	"<input type='radio' onclick='return AddDataText()' id='texta' value=0 name='test'>"+
	"<label for='texta' >Text</label><label id='TextInfo'><i class='fa fa-question-circle-o' style='opacity: 0.35;' aria-hidden='true'></i></label></div>"+
	"<div class='three columns'>"+
	"<input type='radio' onclick='return AddDataMemo()' id='memoa' value=0 name='test'>"+
	"<label for='memoa' >Memo</label><label id='MemoInfo'><i class='fa fa-question-circle-o' style='opacity: 0.35;' aria-hidden='true'></i></label></div>"+
	"<div class='three columns'>"+
	"<input type='radio' onclick='return AddDataInt()' id='inta' value=0 name='test'>"+
	"<label for='inta' >Integer</label><label id='IntInfo'><i class='fa fa-question-circle-o' style='opacity: 0.35;' aria-hidden='true'></i></label></div>"+
	"<div class='three columns'>"+
	"<input type='radio' onclick='return AddDataDouble()' id='doublea' value=0 name='test'>"+
	"<label for='doublea' >Double</label><label id='DoubleInfo'><i class='fa fa-question-circle-o' style='opacity: 0.35;' aria-hidden='true'></i></label></div>"+
	"<div class='two columns'>"+
	"<input type='radio' id='curra' onclick='return AddDataCurr()' value=0 name='test'>"+
	"<label for='curra'>Currency</label><label id='CurrInfo'><i class='fa fa-question-circle-o' style='opacity: 0.35;' aria-hidden='true'></i></label></div>"+
	"<div class='three columns'>"+
	"<input type='radio' id='datea' onclick='return AddDataDate()'value=0 name='test'>"+
	"<label for='datea'>Date/Time</label><label id='DateInfo'><i class='fa fa-question-circle-o' style='opacity: 0.35;' aria-hidden='true'></i></label></div>"+	
	"<div class='three columns'>"+
	"<input type='radio' id='checka' onclick='return AddDataCheck()'value=0 name='test'>"+
	"<label for='checka'>Check-box</label><label id='CheckInfo'><i class='fa fa-question-circle-o' style='opacity: 0.35;' aria-hidden='true'></i></label></div>"+	
	"<div class='three columns'>"+
	"<input type='radio' id='dropa' onclick='return AddDropCheck()' value=0 name='test'>"+
	"<label for='dropa'>Drop Down</label><label id='DropInfo'><i class='fa fa-question-circle-o' style='opacity: 0.35;' aria-hidden='true'></i></label></div>"+	
	"<div class='five columns'>"+
	"<input type='radio' id='hypera' onclick='return AddHyperCheck()' value=0 name='test'>"+
	"<label for='hypera'>Hyperlink</label><label id='HyperFillInfo'><i class='fa fa-question-circle-o' style='opacity: 0.35;' aria-hidden='true'></i></label></div>"+	
	"<div class='four columns'>"+
	"<input type='radio' id='autofa' onclick='return AddAFillCheck()' value=0 name='test'>"+
	"<label for='autofa'>Autofill Box</label><label id='AFillInfo'><i class='fa fa-question-circle-o' style='opacity: 0.35;' aria-hidden='true'></i></label></div>"+	
	"</div>";
	if(alter == null) 
	{
			getp4f.innerHTML +="<div class='twelve columns$'>"+
			"<ul class='actions'>"+
				"<li class='six columns'><input type='submit' onclick= 'return checkCNT()' value='Add Table' class='special' /></li>"+
				"<li class='six columns'><input type='reset' onclick='return EditCreateC()' value='Reset Form' /></li>"+
			"</ul>"+
			"</div></div>"+
			"</form></div>";
	}	
	else{
			getp4f.innerHTML +="<div class='twelve columns$'>"+
			"<ul class='actions'>"+
				"<li class='six columns'><input type='submit' onclick='return checkANC()' value='Add Table' class='special' /></li>"+
				"<li class='six columns'><input type='reset' onclick='return EditCreateC(1)' value='Reset Form' /></li>"+
			"</ul>"+
			"</div></div>"+
			"</form></div>";
	}	
	
	//hovercards in action 
	var TextAHover= 'Use for text or combinations of text and numbers. 255 characters maximum. Will Automatically Use itself as Autofill Source.';
	var MemoAHover= "Memo is used for larger amounts of text. Stores up to 65,536 characters.";
	var IntegerAHover= 'Allows whole numbers between -2,147,483,648 and 2,147,483,647.';
	var DoubleAHover = 'A decimal number, or just decimal, refers to any number written in decimal notation, although it is more commonly used to refer to numbers that have a fractional part separated from the integer part with a decimal separator (e.g. 11.25).';
    var CurrAHover = "Used for currency. Holds up to 15 digits of whole dollars, plus 4 decimal places. Tip: You can choose which country's currency to use."
	var DateAHover = "Used for dates and times."
	var CheckAHover = "A logical field can be displayed as True/False. The True and False Check box can also be used for On/Off, or Yes/No requirements."
	var DropDown = "A drop-down menu is a graphical control element, similar to a list box, that allows the user to choose one value from a list."
	var HyperFillInfo = "In computing, a hyperlink is a reference to data that the reader can directly follow either by clicking, tapping or hovering. A hyperlink points to a whole document."
	var AFillInfo = "Enables users to quickly find and select from a pre-populated list of values as they type, leveraging searching and filtering. Will Automatically Use itself as Autofill Source."
	$("#HyperFillInfo").hovercard({
        detailsHTML: HyperFillInfo
    });
	$("#AFillInfo").hovercard({
        detailsHTML: AFillInfo 
    });
	$("#TextInfo").hovercard({
        detailsHTML: TextAHover
    });
	$("#MemoInfo").hovercard({
        detailsHTML: MemoAHover
    });
	$("#IntInfo").hovercard({
        detailsHTML: IntegerAHover
    });
	$("#DoubleInfo").hovercard({
        detailsHTML: DoubleAHover,
        openOnTop: true,
		openOnLeft: true
	});
	$("#CurrInfo").hovercard({
        detailsHTML:  CurrAHover,
        openOnTop: true
    });
	$("#DateInfo").hovercard({
        detailsHTML:  DateAHover
    });
	$("#CheckInfo").hovercard({
        detailsHTML:  CheckAHover,
        openOnTop: true
    });
	$("#DropInfo").hovercard({
        detailsHTML:  DropDown,
        openOnTop: true,
		openOnLeft: true
    });
};
//adds the required various forms data for the various current data types
function AddDataText()
{
	var counter = document.getElementById("texta");
	if(counter.value >= 1)
		counter.value++;
	else
		counter.value = 1;
	var totalcnt = document.getElementById("totccnt");
	if(totalcnt.value >= 1)
	{
		var getfillarea = document.getElementById("addbox0"+totalcnt.value+"");
		totalcnt.value++;
	}
	else
	{
		var getfillarea = document.getElementById("addbox");
		totalcnt.value = 1;
	}
	getfillarea.innerHTML ="<div class='twelve columns$'>"+
		"<label for='demo-name'>Text Column Name?</label>"+
		"<input type='hidden' id='DataTextO"+counter.value+"' value='' />"+
		"<input name='' id='DataText"+counter.value+"' value='' type='text'/>"+
	"</div><div id='addbox0"+totalcnt.value+"' title='Text'></div>";
	return counter.value;
};
function AddDataMemo()
{
	var counter = document.getElementById("memoa");
	if(counter.value >= 1)
		counter.value++;
	else
		counter.value = 1;
	var totalcnt = document.getElementById("totccnt");
	if(totalcnt.value >= 1)
	{
		var getfillarea = document.getElementById("addbox0"+totalcnt.value+"");
		totalcnt.value++;
	}
	else
	{
		var getfillarea = document.getElementById("addbox");
		totalcnt.value = 1;
	}
	getfillarea.innerHTML ="<div class='twelve columns$'>"+
		"<label for='demo-name'>Memo Column Name?</label>"+
		"<input type='hidden' id='DataMemoO"+counter.value+"' value='' />"+
		"<input name='' id='DataMemo"+counter.value+"' value='' type='text'/>"+
	"</div><div id='addbox0"+totalcnt.value+"' title='Memo'></div>";
	return counter.value;
};
function AddDataInt()
{
	var counter = document.getElementById("inta");
	if(counter.value >= 1)
		counter.value++;
	else
		counter.value = 1;
	var totalcnt = document.getElementById("totccnt");
	if(totalcnt.value >= 1)
	{
		var getfillarea = document.getElementById("addbox0"+totalcnt.value+"");
		totalcnt.value++;
	}
	else
	{
		var getfillarea = document.getElementById("addbox");
		totalcnt.value = 1;
	}
	getfillarea.innerHTML ="<div class='twelve columns$'>"+
		"<label for='demo-name'>Integer Column Name?</label>"+
		"<input type='hidden' id='DataIntO"+counter.value+"' value='' />"+
		"<input name='' id='DataInt"+counter.value+"' value='' type='text'/>"+
		"</div><div id='addbox0"+totalcnt.value+"' title='Int'></div>";
	return counter.value;
};
function AddDataDouble()
{
	var counter = document.getElementById("doublea");
	if(counter.value >= 1)
		counter.value++;
	else
		counter.value = 1;
	var totalcnt = document.getElementById("totccnt");
	if(totalcnt.value >= 1)
	{
		var getfillarea = document.getElementById("addbox0"+totalcnt.value+"");
		totalcnt.value++;
	}
	else
	{
		var getfillarea = document.getElementById("addbox");
		totalcnt.value = 1;
	}
	getfillarea.innerHTML ="<div class='twelve columns$'>"+
		"<label for='demo-name'>Double Column Name?</label>"+
		"<input type='hidden' id='DataDoubleO"+counter.value+"' value='' />"+
		"<input name='' id='DataDouble"+counter.value+"' value='' type='text'/>"+
		"</div><div id='addbox0"+totalcnt.value+"' title='Double'></div>";
	return counter.value;
};
function AddDataCurr()
{
	var counter = document.getElementById("curra");
	if(counter.value >= 1)
		counter.value++;
	else
		counter.value = 1;
	var totalcnt = document.getElementById("totccnt");
	if(totalcnt.value >= 1)
	{
		var getfillarea = document.getElementById("addbox0"+totalcnt.value+"");
		totalcnt.value++;
	}
	else
	{
		var getfillarea = document.getElementById("addbox");
		totalcnt.value = 1;
	}
	getfillarea.innerHTML ="<div class='twelve columns$'>"+
		"<label for='demo-name'>Currency Column Name?</label>"+
		"<input type='hidden' id='DataCurrO"+counter.value+"' value='' />"+
		"<input name='' id='DataCurr"+counter.value+"' value='' type='text'/>"+
		"</div><div id='addbox0"+totalcnt.value+"' title='Curr'></div>";
	return counter.value;
};
function AddDataDate()
{
	var counter = document.getElementById("datea");
	if(counter.value >= 1)
		counter.value++;
	else
		counter.value = 1;
	var totalcnt = document.getElementById("totccnt");
	if(totalcnt.value >= 1)
	{
		var getfillarea = document.getElementById("addbox0"+totalcnt.value+"");
		totalcnt.value++;
	}
	else
	{
		var getfillarea = document.getElementById("addbox");
		totalcnt.value = 1;
	}
	getfillarea.innerHTML ="<div class='twelve columns$'>"+
		"<label for='demo-name'>Date/Time Column Name?</label>"+
		"<input type='hidden' id='DataDateO"+counter.value+"' value='' />"+
		"<input name='' id='DataDate"+counter.value+"' value='' type='text'/>"+
		"</div><div id='addbox0"+totalcnt.value+"' title='Date'></div>";
	return counter.value;
};
function AddDataCheck()
{
	var counter = document.getElementById("checka");
	if(counter.value >= 1)
		counter.value++;
	else
		counter.value = 1;
	var totalcnt = document.getElementById("totccnt");
	if(totalcnt.value >= 1)
	{
		var getfillarea = document.getElementById("addbox0"+totalcnt.value+"");
		totalcnt.value++;
	}
	else
	{
		var getfillarea = document.getElementById("addbox");
		totalcnt.value = 1;
	}
	getfillarea.innerHTML ="<div class='twelve columns$'>"+
		"<label for='demo-name'>Check-box Column Name?</label>"+
		"<input type='hidden' id='DataCheckO"+counter.value+"' value='' />"+
		"<input name='' id='DataCheck"+counter.value+"' value='' type='text'/>"+
		"</div><div id='addbox0"+totalcnt.value+"' title='Check'></div>";
	return counter.value;
};
function AddHyperCheck()
{
	var counter = document.getElementById("hypera");
	if(counter.value >= 1)
		counter.value++;
	else
		counter.value = 1;
	
	var totalcnt = document.getElementById("totccnt");
	if(totalcnt.value >= 1)
	{
		var getfillarea = document.getElementById("addbox0"+totalcnt.value+"");
		totalcnt.value++;
	}
	else
	{
		var getfillarea = document.getElementById("addbox");
		totalcnt.value = 1;
	}
	getfillarea.innerHTML ="<div class='twelve columns$'>"+
		"<label for='demo-name'>Hyperlink Column Name?</label>"+
		"<input type='hidden' id='DataHyperO"+counter.value+"' value='' />"+
		"<input name='' id='DataHyper"+counter.value+"' value='' type='text'/>"+
		"</div><div id='addbox0"+totalcnt.value+"' title='Hyper'></div>";
	return counter.value;
};
function gettablenames()
{
	try{
		var adoConn = new ActiveXObject("ADODB.Connection");
		var adoRS = new ActiveXObject("ADODB.Recordset");
		adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
		adoRS.Open("Select * From TableList Where Obsolete='F'", adoConn);
		var temp = false;
		if(!adoRS.bof)
		{
			adoRS.MoveFirst;
			var temp = "<option value=0></option>";
			
			while(!adoRS.eof)
			{
				temp += "<option value="+
				adoRS.fields("ID")+
				">"+adoRS.fields("TableName")+"</options>";
				adoRS.movenext;
			}
		}
		adoRS.Close();
		adoConn.Close();
		return temp;
	}
	catch(e)
	{
		var e1 = "fillformdata.js had an error in thegettablenames() function";
		ErrorToFile(e, e1);
		location.reload();
	}
}
function AddAFillCheck()
{
	var tablenames = gettablenames();
	var getfillarea = document.getElementById("addautofill");
	var counter = document.getElementById("autofa");
	if(counter.value >= 1)
		counter.value++;
	else
		counter.value = 1;
	var totalcnt = document.getElementById("totccnt");
	if(totalcnt.value >= 1)
	{
		var getfillarea = document.getElementById("addbox0"+totalcnt.value+"");
		totalcnt.value++;
	}
	else
	{
		var getfillarea = document.getElementById("addbox");
		totalcnt.value = 1;
	}
	getfillarea.innerHTML ="<div class='twelve columns'><div class='one-half column'>"+
		"<label for='demo-name'>Source Table Name</label>"+
		"<div class='select-style'><select id='SelATDrop"+counter.value+"' onchange='dropdownmeu("+counter.value+", 1)'>"+
		tablenames +
	"</select></div></div><div class='six columns'>"+
		"<label for='demo-name'>Field That Populates the Autofill Box</label>"+
		"<div class='select-style'><select id='SelAFDrop"+counter.value+"' onchange=''>"+
	"</select></div></div><div class='twelve columns$'>"+
		"<label for='demo-name'>What do you want to call your Autofill box?</label>"+
		"<input type='hidden' id='AutoNameO"+counter.value+"' value='' />"+
		"<input name='' id='AutoName"+counter.value+"' value='' type='text'/>"+
		"</div><div id='addbox0"+totalcnt.value+"' title='AFill'></div>";
	return counter.value;
	
};
function AddDropCheck()
{
	var tablenames = gettablenames();
	var getfillarea = document.getElementById("addndropbox");
	var counter = document.getElementById("dropa");
	if(counter.value >= 1)
		counter.value++;
	else
		counter.value = 1;
	var totalcnt = document.getElementById("totccnt");
	if(totalcnt.value >= 1)
	{
		var getfillarea = document.getElementById("addbox0"+totalcnt.value+"");
		totalcnt.value++;
	}
	else
	{
		var getfillarea = document.getElementById("addbox");
		totalcnt.value = 1;
	}
	getfillarea.innerHTML ="<div class='twelve columns'><div class='one-half column'>"+
		"<label for='demo-name'>Source Table Name</label>"+
		"<div class='select-style'><select id='SelTDrop"+counter.value+"' onchange='dropdownmeu("+counter.value+")'>"+
		tablenames +
	"</select></div></div><div class='six columns'>"+
		"<label for='demo-name'>Field That Populates the DropDown</label>"+
		"<div class='select-style'><select id='SelFDrop"+counter.value+"' onchange=''>"+
	"</select></div></div><div class='twelve columns$'>"+
		"<label for='demo-name'>What do you want to call your Drop Down?</label>"+
		"<input type='hidden' id='SDropNameO"+counter.value+"' value='' />"+
		"<input name='' id='SDropName"+counter.value+"' value='' type='text'/>"+
		"</div><div id='addbox0"+totalcnt.value+"' title='Drop'></div>";
	return counter.value;
	
};
//makes the form to change the table name or a few other options
function TableName(edit)
{
	var temp = returnoptions4CT();
	if(temp != false)
	{
		var getp4f = document.getElementById("fillp4f");
		getp4f.innerHTML = "<a id='close' href='#' class='close'>&nbsp;&nbsp;&nbsp;</a><div id='loge' class='errormessage'></div>"+
		"<label for='demo-name' class='Title4Form'>Table Editor</label>"+
		"<div id='loge' class='errormessage'></div>";
		if(edit == 0)
		{
			getp4f.innerHTML += "<label for='demo-name'>Plese Select A Table You Want to Change the Name For</label>"+
			"<div class='select-style'><select id='OTableName' onchange='return TableNameC()'>"+temp+
			"</select></div><div id='fillp4f2'></div>";
		}
		else if(edit == 1)
		{
			getp4f.innerHTML += "<div class='twelve columns$'>"+
				"<label for='demo-name'>What Table Do you want to change the Column Names</label>"+
				"<div class='select-style'><select id='OTableName' onchange='return ChangeColumnNames()'>"+temp+
			"</select></div><div id='fillp4f3'></div>";
		}
		else if(edit == 2)
		{
			getp4f.innerHTML += "<div class='twelve columns$'>"+
				"<label for='demo-name'>Plese Select A Table You Wish to Add Column's to</label>"+
				"<div class='select-style'><select id='OTableName' onchange='return EditCreateC(1)'>"+temp+
			"</select></div><div id='fillp4f3'></div>";
		}
	}
	else{
		TableCreate();
	}
};
//autofill check if they want to add or remove autofill
function addremautofill()
{
	var getp4f = document.getElementById("fillp4f");
	getp4f.innerHTML = "<div id='loge' class='errormessage'></div>"+
		"<label for='demo-name' class='Title4Form'>Add / Remove Autofill</label>"+
		"<div id='loge' class='errormessage'></div><div id='fillp4f2'>"+
		"<ul class='actions'>"+
			"<li class='six columns'><input type='submit' onclick='return ADDREMAF(1)' value='Add Autofill' class='special' /></li>"+
			"<li class='six columns'><input type='reset' onclick='return ADDREMAF(2)' value='Remove Autofill' /></li>"+
		"</ul></div><a id='close' href='#' class='close'>&nbsp;&nbsp;&nbsp;</a>";
}
//fills up the used field names  when you are looking to remove an autofill
function getusedTN(type)
{
	var adoConn = new ActiveXObject("ADODB.Connection");
	var adoRS = new ActiveXObject("ADODB.Recordset");
	adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
	if(type == 1)
	{
		var filldrop = document.getElementById("AddFTN");
		adoRS.Open("Select Distinct GetTableName From Lookuptable Where Type='A'", adoConn);
	}
	else if(type == 2)
	{
		var getdd = document.getElementById("AddFTN");
		var GTN = getdd.options[getdd.selectedIndex].text;
		adoRS.Open("Select Distinct GetFieldName From Lookuptable Where GetTableName='"+GTN+"' AND Type='A'", adoConn);
	}
	else if(type == 3)
	{
		var filldrop = document.getElementById("AddTTN");
		adoRS.Open("Select Distinct RecTableName From Lookuptable Where Type='A'", adoConn);
	}
	else if(type == 4)
	{
		var getdd = document.getElementById("AddTTN");
		var filldrop = document.getElementById("AddTFN");
		var GTN = getdd.options[getdd.selectedIndex].text;
		adoRS.Open("Select Distinct RecFieldName From Lookuptable Where RecTableName='"+GTN+"' AND Type='A'", adoConn);
	}
	if(!adoRS.bof)
	{
		adoRS.MoveFirst;
		filldrop.options.length = 0;
		var optn = document.createElement("OPTION");
		optn.text = "";
		filldrop.options.add(optn)
		while(!adoRS.eof)
		{
			
			optn = document.createElement("OPTION");
			optn.text = adoRS.fields(0);
			filldrop.options.add(optn)
			adoRS.movenext;
		}
	}
	
	adoRS.Close();
	adoConn.Close();
}
//finishes adding the fill for either add / remove autofill
function ADDREMAF(type)
{
	
	var tablenames = gettablenames();
	var getp4f = document.getElementById("fillp4f2");
	getp4f.innerHTML = "<div class='one-half column'>"+
		"<label for='demo-name'>Select A Table</label>"+
		"<div class='select-style'><select id='AddTTN' onchange='dropdownmeu(0,2)'>"+tablenames+
	"</select></div></div><div class='six columns'>"+
		"<label for='demo-name'>Field of Table to be Autofilled</label>"+
		"<div class='select-style'><select id='AddTFN' onchange=''>"+
	"</select></div></div><div class='one-half column'>"+
		"<label for='demo-name'>Source Table Name</label>"+
		"<div class='select-style'><select id='AddFTN' onchange='dropdownmeu(0,3)'>"+tablenames+
	"</select></div></div><div class='six columns'>"+
		"<label for='demo-name'>Field That Populates the DropDown</label>"+
		"<div class='select-style'><select id='AddFFN' onchange=''>"+
	"</select></div></div>";
	if(type == 2)
	{
		getusedTN(3);
		getusedTN(1);
		document.getElementById("AddTTN").onchange = function(){getusedTN(4);};
		document.getElementById("AddFTN").onchange = function(){getusedTN(2);};
		getp4f.innerHTML += "<ul class='actions'>"+
			"<li class='six columns'><input type='submit' onclick='AutoFillQuery(2)' value='Delete Autofill' class='special' /></li>"+
			"<li class='six columns'><input type='reset' onclick='ADDREMAF(2)' value='Reset' /></li>"+
	"</ul>";
	}
	else
	{
		getp4f.innerHTML += "<ul class='actions'>"+
			"<li class='six columns'><input type='submit' onclick='AutoFillQuery(1)' value='Add Autofill' class='special' /></li>"+
			"<li class='six columns'><input type='reset' onclick='ADDREMAF(1)' value='Reset' /></li>"+
	"</ul>";
	}
}
//delete a table
function TableDelete()
{
	var temp = returnoptions4CT();
	if(temp != false)
	{
		var getp4f = document.getElementById("fillp4f");
		getp4f.innerHTML = "<div id='loge' class='errormessage'></div>"+
		"<label for='demo-name' class='Title4Form'>What Table Do You Want Change The Name For</label>"+
		"<div id='loge' class='errormessage'></div>"+
			"<label for='demo-name'>Plese Select A Process</label>"+
			"<div class='select-style'><select id='ptsel' onchange='return TableDeleteD()'>"+
			temp +
		"</select></div><div id='fillp4f2'></div><a id='close' href='#' class='close'>&nbsp;&nbsp;&nbsp;</a>";
	}
	else{
		TableCreate();
	}
};
//asks the user what table they want to hide or discontinue
function TableObs()
{
	var temp = returnoptions4CT();
	if(temp != false)
	{
		var getp4f = document.getElementById("fillp4f");
		getp4f.innerHTML = "<div id='loge' class='errormessage'></div>"+
		"<label for='demo-name' class='Title4Form'>What Table Do You Want to make Obsolete</label>"+
		"<div id='loge' class='errormessage'></div>"+
			"<label for='demo-name'>Plese Select A Process</label>"+
			"<div class='select-style'><select id='ptsel' onchange='return TableObsD()'>"+
			temp +
		"</select><div id='fillp4f2'></div><a id='close' href='#' class='close'>&nbsp;&nbsp;&nbsp;</a>";
	}
	else{
		TableCreate();
	}
};
//Gets the User to Enter in the New Table Name
function TableNameC()
{
	var OTableName = document.getElementById("OTableName");
	
	var getp4f = document.getElementById("fillp4f2");
	getp4f.innerHTML ="<form method='post' action='#' >"+
		"<label for='demo-name'>What do you want to Rename your table To</label>"+
		"<input name='"+OTableName.options[OTableName.selectedIndex].text+"' id='tableName' value='' type='text'/>"+
	"<div id='fillp4f3'>"+
		"<ul class='actions'>"+
			"<li class='six columns'><input type='submit' onclick='return ChangeTN()' value='Change Table Name' class='special' /></li>"+
			"<li class='six columns'><input type='reset' value='Reset Form' /></li>"+
		"</ul></div></form>";
};
//Asks The User if they want to delete a Table
function TableDeleteD()
{
	var ptsel = document.getElementById("ptsel");
	var getp4f = document.getElementById("fillp4f2");
	getp4f.innerHTML ="<form method='post' action='#' >"+
		"<input name='"+ptsel.options[ptsel.selectedIndex].text+"' id='tableName' value='' type='hidden'/>"+
	"<div id='fillp4f3'>"+
		"<ul class='actions'>"+
			"<li class='six columns'><input type='submit' onclick='return DeleteTN()' value='Delete Table' class='special' /></li>"+
			"<li class='six columns'><input type='reset' value='Reset Form' /></li>"+
		"</ul></div>"+
	"</form>";
};
//Disc table
function TableObsD()
{
	var ptsel = document.getElementById("ptsel");
	var getp4f = document.getElementById("fillp4f2");
	getp4f.innerHTML ="<form method='post' action='#' >"+
		"<input name='"+ptsel.options[ptsel.selectedIndex].text+"' id='tableName' value='' type='hidden'/>"+
	"<div id='fillp4f3'>"+
		"<ul class='actions'>"+
			"<li class='six columns'><input type='submit' onclick='return DiscontinueTN()' value='Discontinue/Enable Table' class='special' /></li>"+
			"<li class='six columns'><input type='reset' value='Reset Form' /></li>"+
		"</ul></div></form>";
};
//link table to another process\
//step 2
function TableLinkA()
{
	
	var ptsel = document.getElementById("ptsel");
	var process = returnoptions4dd(ptsel.options[ptsel.selectedIndex].value);
	var getp4f = document.getElementById("fillp4f2");
	getp4f.innerHTML ="<form method='post' action='#' >"+
		"<input name='"+ptsel.options[ptsel.selectedIndex].text+"' id='tableName' value='' type='hidden'/>"+
			"<label for='demo-name'>Plese Select A Table</label>"+
			"<div class='select-style'><select id='linkto' >"+
			process +
		"</select>"+
	"<div id='fillp4f3'>"+
		"<ul class='actions'>"+
			"<li class='six columns'><input type='submit' onclick='return TableLinking()' value='Link Table' class='special' /></li>"+
			"<li class='six columns'><input type='reset' value='Reset Form' /></li>"+
		"</ul></div></div></form>";
};
//link table to another process if table exists otherwise goes to step two
function TableLink()
{
	
	var temp = returnoptions4CT();
	if(temp != false)
	{
		var getp4f = document.getElementById("fillp4f");
		getp4f.innerHTML = "<div id='loge' class='errormessage'></div>"+
		"<label for='demo-name' class='Title4Form'>What Table do you want to Link to Another Process</label>"+
		"<div id='loge' class='errormessage'></div>"+
			"<label for='demo-name'>Plese Select A Table</label>"+
			"<div class='select-style'><select id='ptsel' onchange='return TableLinkA()'>"+temp+
		"</select><div id='fillp4f2'></div><a id='close' href='#' class='close'>&nbsp;&nbsp;&nbsp;</a>";
	}
	else{
		TableCreate();
	}
};
//adds admin to a table or enable Delete
function AddAdminOEn()
{
	var getp4f = document.getElementById("fillp4f");
	getp4f.innerHTML = "<div id='loge' class='errormessage'></div>"+
	"<label for='demo-name' class='Title4Form'>Form for Add Administrator / Toggle Ability to Delete</label>"+
	"<div id='loge' class='errormessage'></div>"+
	"<form method='post'>"+
		"<label for='demo-name'>Name of New Administrator</label>"+
		"<input type='text' name='demo-name' id='NAdName' value='' />"+
	"</div><ul class='actions'>"+
			"<li class='four columns'><input type='submit' value='Add Administrator' onclick='return AddAdmin()' class='special' /></li>"+
			"<li class='four columns'><input type='submit' onclick='return toggledelete()' value='Toggle Ability to Delete' class='special' /></li>"+
			"<li class='four columns'><input type='reset' value='Reset Form' /></li>"+
		"</ul></div></form><a id='close' href='#' class='close'>&nbsp;&nbsp;&nbsp;</a>";
};
//saves an attachment
function SaveAttachment()
{
	var getp4f = document.getElementById("fillp4f");
	getp4f.innerHTML = "<div id='loge' class='errormessage'></div>"+
	"<label for='demo-name' class='Title4Form'>Form for Add Attachment</label>"+
	"<div id='loge' class='errormessage'></div>"+
	"<form method='post' action='#' onsubmit='return addAttachment()'>"+
		"<label for='demo-name'>Attachment Name</label>"+
		"<input type='text' name='demo-name' id='atttitle' value='' />"+
		"<label for='demo-message'>Attachment</label><label class='fileContainer'>Click Here to Upload File And Create Hyperlink <input type='file' id='Attachment4MP' onchange=''/></label>"+
		"</div><ul class='actions'>"+
			"<li class='six columns'><input type='submit' value='Add Attachment' class='special' /></li>"+
			"<li class='six columns'><input type='reset' value='Reset Form' /></li>"+
		"</ul></div></form><a id='close' href='#' class='close'>&nbsp;&nbsp;&nbsp;</a>";
}