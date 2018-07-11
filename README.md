# Developer

 public class UploadControlHelperClockTransaction
        {

            public static List<DataRow> ListFile;
            string[] FileExtensions = new string[] { ".DOC", ".TXT", ".PDF", "xlsx", "XLS", "CSV" };

            public static readonly UploadControlValidationSettings ValidationSettingsFile = new UploadControlValidationSettings
            {
                AllowedFileExtensions = new string[] { ".DOC", ".TXT", ".PDF" },
                MaxFileSize = 4000000,
            };

            public static readonly UploadControlValidationSettings ValidationSettings = new UploadControlValidationSettings
            {
                MaxFileSize = 4000000,
            };


            public static void FileUploadComplete(object sender, FileUploadCompleteEventArgs e)
            {
                if (e.UploadedFile.IsValid)
                {
                    // string[] ImageExtensions = new string[] { ".PNG", ".JPG", ".GIF", ".JPEG" };
                    string FilePath = "";
                    string FileName = "";

                    FileName = System.IO.Path.GetFileNameWithoutExtension(e.UploadedFile.FileName);

                    var url = System.Web.HttpContext.Current.Request.Url.AbsolutePath;
                    var array = url.Split(new[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
                    string path = "Areas\\HumanResource\\Views\\ClockTransaction\\TempClockTransaction";
                    string root = System.Web.HttpContext.Current.Server.MapPath("~");
                    string folder = root + path;
                    DataTable dt;
                    if (FileName.Length > 0)
                        FilePath = folder + "/" + FileName + Path.GetExtension(e.UploadedFile.FileName);
                    else
                        FilePath = folder + "/" + e.UploadedFile.FileName;

                    if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
                    e.UploadedFile.SaveAs(FilePath);


                    IUrlResolutionService urlResolver = sender as IUrlResolutionService;
                    if (urlResolver != null)
                        e.CallbackData = urlResolver.ResolveClientUrl(FilePath);
                    string extension = System.IO.Path.GetExtension(FilePath).ToLower();
                    if (extension == ".csv")
                    {
                        dt = UtilityConvertDataTable.ConvertCSVtoDataTable(FilePath);
                        List<DataRow> list = new List<DataRow>();
                        foreach (DataRow dr in dt.Rows)
                        {
                            list.Add(dr);
                        }
                        ListFile = list;
                    }
                    else if (extension.Trim() == ".xls")
                    {
                        dt = UtilityConvertDataTable.ConvertXSLXtoDataTable(FilePath);
                        List<DataRow> list = new List<DataRow>();
                        foreach (DataRow dr in dt.Rows)
                        {
                            list.Add(dr);
                        }
                        ListFile = list;
                    }
                    else if (extension.Trim() == ".xlsx")
                    {
                        dt = UtilityConvertDataTable.ConvertXSLXtoDataTable(FilePath);
                        List<DataRow> list = new List<DataRow>();
                        foreach (DataRow dr in dt.Rows)
                        {
                            list.Add(dr);
                        }
                        ListFile = list;

                    }


                }
            }


        }
        public static class UtilityConvertDataTable
        {
            public static DataTable ConvertCSVtoDataTable(string strFilePath)
            {
                Workbook workbook = new Workbook();
                workbook.LoadDocument(strFilePath, DocumentFormat.Csv);               
                Worksheet workSheet = workbook.Worksheets[0];
                Range usedRange = workSheet.GetUsedRange();               
                DataTable dataTable = workSheet.CreateDataTable(usedRange.CurrentRegion, true, true);                
                for (int i = 1; i <= usedRange.RowCount - 1; i++)
                {
                    DataRow newRow = dataTable.NewRow();
                    for (int j = 0; j <= usedRange.CurrentRegion.ColumnCount - 1; j++)
                    {
                        newRow[j] = workSheet.Cells[i, j].DisplayText;
                    }
                    dataTable.Rows.Add(newRow);
                }
                return dataTable;



            }

            public static DataTable ConvertXSLXtoDataTable(string strFilePath)
            {
                Workbook workbook = new Workbook();
                string extension = System.IO.Path.GetExtension(strFilePath).ToLower();
                if (extension.Trim() == ".xls")
                    workbook.LoadDocument(strFilePath, DocumentFormat.Xls);
                else
                    workbook.LoadDocument(strFilePath, DocumentFormat.Xlsx);
                Worksheet workSheet = workbook.Worksheets[0];
                Range usedRange = workSheet.GetUsedRange();
                DataTable dataTable = workSheet.CreateDataTable(usedRange.CurrentRegion, true, true);

                for (int i = 1; i <= usedRange.RowCount - 1; i++)
                {
                    DataRow newRow = dataTable.NewRow();
                    for (int j = 0; j <= usedRange.CurrentRegion.ColumnCount - 1; j++)
                    {
                        newRow[j] = workSheet.Cells[i, j].DisplayText;
                    }
                    dataTable.Rows.Add(newRow);
                }
                return dataTable;

            }
        }
        
        
        ////////////////////
           [HttpPost]
        public ActionResult AttachUploadClockTransaction()

        {

            UploadedFile[] file = UploadControlExtension.GetUploadedFiles("uploadFile", UploadControlHelperClockTransaction.ValidationSettings, UploadControlHelperClockTransaction.FileUploadComplete);
            //return Json(new { Result = "" }, JsonRequestBehavior.AllowGet);
            return null;

        }
        ///////////////
        
function FileUploadComplete(s, e) {
     array1 = new Array();
    var _SeperateType = SeperateType.GetValue();
    ArrayGrid();
    array1;
    if (e.callbackData) {
        $.ajax({
            url: '/Operation/InvStockTaking/FileComplete',
            type: 'POST',
            dataType: 'HTML', 
            //dataType: 'json',           
            data: {
                ObjectVWItemStock: array1 ,
                SeperateType: _SeperateType
            },

            success: function (data) {

                
                $('#gridg').html(data);




            }
        });

    }
   
}


function DisplayUpload() {

    $("#Upload").show();
}

var array1 = new Array();
function ArrayGrid() {

    
    var leng = GridDetailsItem.GetVisibleRowsOnPage();
    for (var i = 0; i < leng; i++) {
        var object = {
            
            ItemID: GridDetailsItem.batchEditApi.GetCellValue(i, "ItemID"),
            UnitID: GridDetailsItem.batchEditApi.GetCellValue(i, "UnitID"),
            Qty: GridDetailsItem.batchEditApi.GetCellValue(i, "Qty"),
            Actual: GridDetailsItem.batchEditApi.GetCellValue(i, "Actual"),
            Differences: GridDetailsItem.batchEditApi.GetCellValue(i, "Differences"),
            
        };

        array1.push(object);
    }
}
///////

     [HttpPost]
        public ActionResult FileComplete(VWItemStock[] ObjectVWItemStock, int SeperateType)
        {
            var listUplaodFile = UploadControlHelperInvStockTaking.ListFile;

            var _StockTakingHdrDto = new List<string>();

            if (SeperateType == (int)HardCodeIDBase.SeparateTypeREF_Slash)
            {
                foreach (DataRow row in listUplaodFile)
                {
                    var RowUpload = row[0].ToString();
                    var first = RowUpload.Split('/').First();//get bar code 

                    if (RowUpload.Contains('/') == true)
                    {
                        var last = RowUpload.Split('/').Last();//get QtyActual
                        _StockTakingHdrDto.Add(last);
                    }
                    
                }
                //var last = x.Split('/').Last();
                
            }
            var model = _IStockTakingService.GetItemStock(1, 1, 1);
            
            RemList<VWItemStock> RemList = model;
          
            return PartialView("_grdStockTakingDetailsBatch", model);
            // return Json(new { Result = _StockTakingHdrDto }, JsonRequestBehavior.AllowGet);

        }

        
        
        
