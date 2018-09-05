// Validates that the input string is a valid date formatted as "mm/dd/yyyy"
function isValidDate(dateString)
{
	try{
    // First check for the pattern
    if(!/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(dateString))
        return false;
    // Parse the date parts to integers
    var parts = dateString.split("/");
    var day = parseInt(parts[1], 10);
    var month = parseInt(parts[0], 10);
    var year = parseInt(parts[2], 10);

    // Check the ranges of month and year
    if(year < 1000 || year > 3000 || month == 0 || month > 12)
        return false;

    var monthLength = [ 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 ];

    // Adjust for leap years
    if(year % 400 == 0 || (year % 100 != 0 && year % 4 == 0))
        monthLength[1] = 29;

    // Check the range of the day
    return day > 0 && day <= monthLength[month - 1];
	}catch(e){ErrorToFile(e);}
};
//checks if number is float maybe all the way up into double territory
var isFloat = function(n) {
	try{
    n = n.length > 0 ? Number(n) : false;
    return (n === parseFloat(n));
	}catch(e){ErrorToFile(e);}
};
//checks if it is a number
var isInteger = function(n) {
	try{
    n = n.length > 0 ? Number(n) : false;
    return (n === parseInt(n));
	}catch(e){ErrorToFile(e);}
};
//checks both int and float situations 
var isNumeric = function(n){
	try{
   if(isInteger(n) || isFloat(n)){
        return true;
   }
   return false;
	}catch(e){ErrorToFile(e);}
};
function createrecord4autofill(ID, Edit, Del)
{
	try{
		var itstxt = document.getElementById("TxtFill");
		var itsMemo = document.getElementById("MemoFill");
		var itsLong = document.getElementById("LongFill");
		var itsDou = document.getElementById("DouFill");
		var itsCurr = document.getElementById("CurrFill");
		var itsDate = document.getElementById("DateFill");
		var itsCheck = document.getElementById("CheckFill");
		var itsDrop = document.getElementById("DropFill");
		var itsHyper = document.getElementById("hyperfill");
		var itsFill = document.getElementById("Auto-Fill");
		var loge = document.getElementById("loge");
		setcolorofloge(0);
		var defaults = false;
		var bool = true;
		
		for(var i = 1; i <= itstxt.name;i++)
		{
			var temp = document.getElementById("TxtFill"+i+"");
			if(temp.value == "")
			{
				if(i == 1)
					loge.innerHTML += "Please Make Sure "+temp.name+" ";
				else
					loge.innerHTML += ", "+temp.name+"";
				if(bool == true)
					temp.focus(), bool = false;
				temp.style.borderColor = "#7e3d3f";
				if(i == itstxt.name)
					loge.innerHTML += " Field(S) has an Entry.<br />";
			}
			else
				temp.style.borderColor = "#476A34";
		}
		for(var i = 1; i <= itsFill.name;i++)
		{
			var temp = document.getElementById("AutoF#"+i+"");
			if(temp.value == "" || temp.value == null)
			{
				if(i == 1)
					loge.innerHTML += "Please Make Sure "+temp.name+" ";
				else
					loge.innerHTML += ", "+temp.name+"";
				if(bool == true)
					temp.focus(), bool = false;
				temp.style.borderColor = "#7e3d3f";
				if(i == itsFill.name)
					loge.innerHTML += " Field(S) has an Entry.<br />";
			}
			else
				temp.style.borderColor = "#476A34";
		}
		for(var i = 1; i <= itsLong.name;i++)
		{
			var temp = document.getElementById("LongFill"+i+"");
			if(!isInteger(temp.value))
			{
				if(i == 1)
					loge.innerHTML += "Please Make Sure "+temp.name+"";
				else
					loge.innerHTML += ", "+temp.name+"";
				if(bool == true)
					temp.focus(), bool = false;
				temp.style.borderColor = "#7e3d3f";
				if(i == itsLong.name)
					loge.innerHTML += "Field(S) is an Integer(IE. 1249).<br />";
			}
			else
				temp.style.borderColor = "#476A34";
		}
		for(var i = 1; i <= itsDou.name;i++)
		{
			var temp = document.getElementById("DouFill"+i+"");
			if(!isNumeric(temp.value))
			{
				if(i == 1)
					loge.innerHTML += "Please Make Sure "+temp.name+"";
				else
					loge.innerHTML += ", "+temp.name+"";
				if(bool == true)
					temp.focus(), bool = false;
				temp.style.borderColor = "#7e3d3f";
				if(i == itsDou.name)
					loge.innerHTML += "Field(S) has a Double(IE. 4.65).<br />";
			}
			else
				temp.style.borderColor = "#476A34";
		}
		for(var i = 1; i <= itsCurr.name;i++)
		{
			var temp = document.getElementById("CurrFill"+i+"");
			if(!isNumeric(temp.value))
			{
				if(i == 1)
					loge.innerHTML += "Please Make Sure "+temp.name+"";
				else
					loge.innerHTML += ", "+temp.name+"";
				if(bool == true)
					temp.focus(), bool = false;
				temp.style.borderColor = "#7e3d3f";
				if(i == itsCurr.name)
					loge.innerHTML += "Field(S) is a Currency(IE. 4.65).<br />";
			}
			else
				temp.style.borderColor = "#476A34";
		}
		for(var i = 1; i <= itsDate.name;i++)
		{
			var temp = document.getElementById("DateFill"+i+"");
			if(!isValidDate(temp.value))
			{
				if(i == 1)
					loge.innerHTML += "Please Make Sure "+temp.name+"";
				else
					loge.innerHTML += ", "+temp.name+"";
				if(bool == true)
					temp.focus(), bool = false;
				temp.style.borderColor = "#7e3d3f";
				if(i == itsDate.name)
					loge.innerHTML += "Field(S) is a Date(IE. MM/DD/YYYY).<br />";
			}
			else
				temp.style.borderColor = "#476A34";
		}
		for(var i = 1; i <= itsDrop.name;i++)
		{
			var temp = document.getElementById("DropFill"+i+"");
			if(temp.options[temp.selectedIndex].text == "")
			{
				if(i == 1)
					loge.innerHTML += "Please Select an Entry for "+temp.name+" ";
				else
					loge.innerHTML += ", "+temp.name+"";
				if(bool == true)
					temp.focus(), bool = false;
				temp.style.backgroundColor = "#7e3d3f";
				if(i == itstxt.name)
					loge.innerHTML += " Field(S).<br />";
			}
			else
				temp.style.backgroundColor = "#476A34";
		}
		if(bool == true)
		{
	//this one doesnt use accessdb.js I needed more fucntionality so I did it myself
			var adoConn = new ActiveXObject("ADODB.Connection");
			var adoRS = new ActiveXObject("ADODB.Recordset");
			var TableName = document.getElementById("ID#"+ID+"").name;
			adoConn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='Abcon.mdb'");
			if(Edit == null && Del == null)
			{
				adoRS.Open("Select * From ["+TableName+"]", adoConn, 1, 3);
				adoRS.AddNew;
			}
			else if(Edit != null && Del == null)
			{
				adoRS.Open("Select * From ["+TableName+"] Where ID="+Edit+"", adoConn, 1, 3);
				adoRS.Edit;
			}
			else
			{
				adoRS.Open("Select * From ["+TableName+"] Where ID="+Del+"", adoConn, 1, 3);
				if(!adoRS.bof)
					adoRS.Delete;
				adoRS.Close();
				adoConn.Close();
				loge.innerHTML += "The Record has been Deleted.";
				loge.style.color = "White";
				loge.style.backgroundColor = "#4c6c96";
				loge.style.borderColor = "#4b2426";
				if(Edit != null || Del != null)
					generateTable(ID);
				return false;
			}
			for(var i = 1; i<= itsCheck.name; i++)
			{
				var temp = document.getElementById("CheckYFill"+i+"");
				if(temp.checked)
					adoRS.Fields(temp.name).value = -1;
				else
					adoRS.Fields(temp.name).value = 0;
					
			}
			for(var i = 1; i<= itsMemo.name; i++)
			{
				var temp = document.getElementById("MemoFill"+i+"");
				if(temp.value != "")
					adoRS.Fields(temp.name).value = temp.value;
				temp.value = "";
					
			}
			for(var i = 1; i<= itsHyper.name; i++)
			{
				var temp = document.getElementById("HyperFill"+i+"");
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
					var finaldir = window.location.href;
					finaldir = finaldir.replace("index.hta", "");
					finaldir = finaldir.replace("adminpanel.hta", "");
					finaldir = finaldir.replace("processviewer.hta", "");
					finaldir += "DatabaseFiles/"+filename+"";
					finaldir = finaldir.replace("file:///","");
					alert(finaldir);
					myObject.CopyFile(origurl, finaldir);
					adoRS.Fields(temp.alt).value = filename;
				}
					
			}
			for(var i = 1; i<= itstxt.name; i++)
			{
				var temp = document.getElementById("TxtFill"+i+"");
				adoRS.Fields(temp.name).value = temp.value;
				temp.value = "";
			}
			for(var i = 1; i<= itsFill.name; i++)
			{
				var temp = document.getElementById("AutoF#"+i+"");
				adoRS.Fields(temp.name).value = temp.value;
				temp.value = "";
			}
			for(var i = 1; i<= itsLong.name; i++)
			{
				var temp = document.getElementById("LongFill"+i+"");
				adoRS.Fields(temp.name).value = temp.value;	
				temp.value = "";
			}
			for(var i = 1; i<= itsDou.name; i++)
			{
				var temp = document.getElementById("DouFill"+i+"");
				adoRS.Fields(temp.name).value = temp.value;	
				temp.value = "";
			}
			for(var i = 1; i<= itsCurr.name; i++)
			{
				var temp = document.getElementById("CurrFill"+i+"");
				adoRS.Fields(temp.name).value = temp.value;	
				temp.value = "";
			}
			for(var i = 1; i<= itsDate.name; i++)
			{
				var temp = document.getElementById("DateFill"+i+"");
				adoRS.Fields(temp.name).value = temp.value;	
				temp.value = "MM/DD/YYYY";
			}
			for(var i = 1; i<= itsDrop.name; i++)
			{
				var temp = document.getElementById("DropFill"+i+"");
				adoRS.Fields(temp.name).value = temp.options[temp.selectedIndex].text;
				temp.selectedIndex = 0;
			}
			adoRS.Update;
			adoRS.Close();
			adoConn.Close();
			setcolorofloge(1)
			loge.innerHTML += "The Record has been Added.";
			if(Edit != null || Del != null)
				generateTable(ID);
		}
		if(Edit != null || Del != null)
			return true;
		return false;	
	}
	catch(e)
	{
		var e1 = "errorcheckPV.js had an error in createrecord4autofill(ID, Edit, Del) function";
		ErrorToFile(e, e1);
		location.reload();
	}
}
function exportselected()
{
	try{
		var mytable = document.getElementById("TheExport");
		var colCount = mytable.getElementsByTagName("tr")[0].getElementsByTagName("th").length;
		var newArray = [];
		for (var j = 0; j < colCount-1; j++)
			newArray.push(document.getElementById("Exportxls"+(j+1)+"").checked);
		var sql = getsqlfromsearch();
		generateTable(getCookie("GTID"), sql);
		var str = "";
		var mytable = document.getElementById("TheExport");
		var rowCount = mytable.rows.length;
		var colCount = mytable.getElementsByTagName("tr")[0].getElementsByTagName("th").length;
		for (var j = 0; j < colCount-1; j++)
			document.getElementById("Exportxls"+(j+1)+"").checked = newArray[j];
		var xls = new ActiveXObject("Excel.Application");
		var newBook = xls.Workbooks.Add;
		newBook.Worksheets(1).Activate;
		newBook.Worksheets(3).delete;
		newBook.Worksheets(2).delete;
		xls.visible = true;
		var TN = getCookie("TableName");
		newBook.Worksheets(1).Name= TN;
		var printrow = 0;
		for (var i = 0; i < rowCount; i++) {
			if(i == 1)
				i++;
			var printcol = 0;
			for (var j = 0; j < colCount-1; j++) {
				if(document.getElementById("Exportxls"+(j+1)+"").checked == true)
				{
					if (i == 0) {
						str = mytable.getElementsByTagName("tr")[i].getElementsByTagName("th")[j].innerText;
					}
					else {
						str = mytable.getElementsByTagName("tr")[i].getElementsByTagName("td")[j].innerText;
					}
					 newBook.Worksheets(1).Cells(printrow + 1, printcol + 1).Value = str;
					if((i+1) == 1)
						 newBook.Worksheets(1).Cells(printrow+1, printcol+1).Interior.ColorIndex="35";
					 newBook.Worksheets(1).Cells (printrow+1, printcol+1). Borders.Weight = 2;
					 newBook.Worksheets(1).Columns(printcol+1).AutoFit
					 newBook.Worksheets(1).Columns(printcol+1).WrapText = true;
					printcol++;
				}
			}
			printrow++;
		}
		if(printcol == 0)
			return false;
		generateTable(getCookie("GTID"));
		var mytable = document.getElementById("TheExport");
		var colCount = mytable.getElementsByTagName("tr")[0].getElementsByTagName("th").length;
		for (var j = 0; j < colCount-1; j++)
			document.getElementById("Exportxls"+(j+1)+"").checked = newArray[j];
	}
	catch(e)
	{
		var e1 = "errorcheckPV.js had an error in exportselected() function";
		ErrorToFile(e, e1);
		location.reload();
	}
}
function selectall()
{
	try{
	var mytable = document.getElementById("TheExport");
	var getcols = mytable.getElementsByTagName("tr")[0].getElementsByTagName("th").length;
	for(var i = 1; i < getcols;i++)
		document.getElementById("Exportxls"+i+"").checked = true;
	exportselected();
	}catch(e){ErrorToFile(e);}
}
function showexport()
{
	try{
	var texport = document.getElementById("exporthidden");
	texport.style.display = 'block';
	}catch(e){ErrorToFile(e);}
};































