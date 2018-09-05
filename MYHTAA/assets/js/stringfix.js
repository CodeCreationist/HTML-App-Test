//Basically tells the user what characters are allowed as column names
//edits the users entry then adds in a modified if it is allowed.
function stringcheck(fixed)
{
	try{
		var bool = false;
		if(fixed.indexOf("[") >= 0)
			bool = true;
		if(fixed.indexOf(".") >= 0)
			bool = true;
		if(fixed.indexOf("]") >= 0)
			bool = true;
		if(fixed.indexOf("!") >= 0)
			bool = true;
		if(fixed.indexOf("`") >= 0)
			bool = true;
		if(fixed.indexOf("?") >= 0)
			bool = true;
		if(fixed.indexOf("=") >= 0)
			bool = true;
		if(bool == false)
			return true;
		else
			return false;
	}catch(e){ErrorToFile(e);}
}
function stringfix(string, option)
{
	try{
		var fixed = string;
		fixed = fixed.split(']').join("");
		fixed = fixed.split('[').join("");
		fixed = fixed.split('.').join("");
		fixed = fixed.split('`').join("");
		fixed = fixed.split('!').join(" ");
		fixed = fixed.split('?').join(" ");
		fixed = fixed.split('=').join("");
		return fixed;
	}catch(e){ErrorToFile(e);}
}