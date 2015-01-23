using CRM.Common;
using CRM.Model;
using NLog;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Hi_Life_API_Dept
{
    class Flow
    {
        private static Logger _logger = LogManager.GetCurrentClassLogger();
        internal void DoFlow()
        {
            DataSyncModel dataSync = new DataSyncModel();
            EnvironmentSetting.LoadSetting();
            if (EnvironmentSetting.ErrorType == ErrorType.None)
            {
                //create dataSync
                dataSync.CreateDataSyncForCRM("套組");
                if (EnvironmentSetting.ErrorType == ErrorType.None)
                {
                    XSSFWorkbook workbook;
                    try
                    {
                        //讀取專案內中的sample.xls 的excel 檔案
                        using (FileStream file = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "套組.xlsx", FileMode.Open, FileAccess.Read))
                        {
                            workbook = new XSSFWorkbook(file);
                        }

                        //讀取Sheet1 工作表
                        var sheet = workbook.GetSheet("套組");

                        Console.WriteLine("連線成功!!");
                        Console.WriteLine("開始執行...");

                        Model model = new Model();
                        model.doModel(sheet.GetRow(0).Cells);
                        model.deleteAll();

                        if (EnvironmentSetting.ErrorType == ErrorType.None)
                        {
                            int success = 0;
                            int fail = 0;
                            int partially = 0;

                            for (int i = 1; i <= sheet.LastRowNum; i++)
                            {
                                if (sheet.GetRow(i) != null) //null is when the row only contains empty cells 
                                {
                                    var row = sheet.GetRow(i);
                                    TransactionStatus transactionStatus;
                                    TransactionType transactionType;

                                    //create product
                                    transactionType = TransactionType.Insert;
                                    transactionStatus = model.CreateForCRM(row);

                                    //create datasyncdetail
                                    if (EnvironmentSetting.ErrorType == ErrorType.None)
                                    {
                                        switch (transactionStatus)
                                        {
                                            case TransactionStatus.Success:
                                                success++;
                                                break;
                                            case TransactionStatus.Fail:
                                                //新增、更新資料有錯誤 則新增一筆detail
                                                dataSync.CreateDataSyncDetailForCRM(row.GetCell(0).NumericCellValue.ToString(), row.GetCell(13) + row.GetCell(15).NumericCellValue.ToString(), transactionType, transactionStatus);
                                                fail++;
                                                break;
                                            default:
                                                fail++;
                                                break;
                                        }

                                        //新增detail錯誤 則結束
                                        if (EnvironmentSetting.ErrorType != ErrorType.None)
                                        {
                                            _logger.Info(EnvironmentSetting.ErrorMsg);
                                            break;
                                        }
                                    }
                                }
                            }
                            //更新DataSync 成功、失敗、完成時間
                            dataSync.UpdateDataSyncForCRM(success, fail, partially);
                            Console.WriteLine("成功 : " + success);
                            Console.WriteLine("失敗 : " + fail);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("檔案讀取錯誤");
                        Console.WriteLine(ex.Message);
                        EnvironmentSetting.ErrorMsg = "欄位讀取錯誤\n" + ex.Message;
                        EnvironmentSetting.ErrorType = ErrorType.EXCEL;
                    }
                }
            }
            switch (EnvironmentSetting.ErrorType)
            {
                case ErrorType.None:
                    break;

                case ErrorType.INI:
                case ErrorType.CRM:
                case ErrorType.DATASYNC:
                    _logger.Info(EnvironmentSetting.ErrorMsg);
                    break;

                case ErrorType.EXCEL:
                case ErrorType.DATASYNCDETAIL:
                    dataSync.UpdateDataSyncWithErrorForCRM(EnvironmentSetting.ErrorMsg);
                    break;

                default:
                    break;
            }
            Console.WriteLine("執行完畢...");
            Console.ReadLine();
        }
    }
}
