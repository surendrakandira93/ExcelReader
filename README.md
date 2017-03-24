# ExcelReader
ExcelReader read excel file and return json , Dictonary, And collecton

How to use

C# code :


string sWebRootFolder = _hostingEnvironment.WebRootPath;
string sFileName = @"excelfile.xlsx";
string path = Path.Combine(sWebRootFolder, sFileName);           
IExcelDataReader iExcelDataReader = ExcelReaderFactory.Reader(path); 
var response = iExcelDataReader.AsIEnumerable(); // return Dictonary
// var response = iExcelDataReader.AsIEnumerable(); // return json

