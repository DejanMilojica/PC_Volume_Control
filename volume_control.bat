@if (@a==@b) @end /*

:: batch portion

@ECHO OFF

cscript /e:jscript "%~f0"


:: JScript portion */

var shl = new ActiveXObject("WScript.Shell");
for (var i=0; i<=20; i++) {
    shl.SendKeys(String.fromCharCode(0xAE));
    WScript.Sleep(2000);
}
