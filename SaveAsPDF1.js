//Javascript to convert word document to pdf 
//For TPA team process improvement
//Developed by Thillai 6 Jan 2017
//Reading a file using filesystemobject
var tpa = new ActiveXObject("Scripting.FileSystemObject");
//Arguments passed via command line can be accessed by WScript where 0 is the argument
var docPath = WScript.Arguments(0);
//To get the pathname
docPath = tpa.GetAbsolutePathName(docPath);
//Using replace string method, with regular expression for search patterns and modifiers
var pdfPath = docPath.replace(/\.doc[^.]*$/, ".pdf");
var objWord = null;

try
{
    WScript.Echo("Saving '" + docPath + "' as '" + pdfPath + "'...");

    objWord = new ActiveXObject("Word.Application");
    objWord.Visible = false;

    var objDoc = objWord.Documents.Open(docPath);

    var wdFormatPdf = 17;
    objDoc.SaveAs(pdfPath, wdFormatPdf);
    objDoc.Close();

    WScript.Echo("Conversion Succesfully Completed.");
}
finally
{
    if (objWord != null)
    {
        objWord.Quit();
    }
}
