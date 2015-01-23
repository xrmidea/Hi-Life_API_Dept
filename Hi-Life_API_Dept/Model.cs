using CRM.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hi_Life_API_Dept
{
    class Model
    {
        private IOrganizationService service = EnvironmentSetting.Service;

        List<Entity> deptList;
        List<Entity> storeInformationList;

        List<OptionMetadata> new_6;
        List<OptionMetadata> new_5;
        List<OptionMetadata> new_17;
        List<OptionMetadata> new_13;
        List<OptionMetadata> new_4;
        List<OptionMetadata> new_14;
        List<OptionMetadata> new_9;
        List<OptionMetadata> new_16;
        List<OptionMetadata> new_12;
        List<OptionMetadata> new_7;
        List<OptionMetadata> new_8;
        List<OptionMetadata> new_11;
        List<OptionMetadata> new_10;

        String[] fieldArray = {"new_stonenumber","new_stonename","null","null",
                                "new_6","new_5","new_3","new_50",
                                "new_17","new_13","new_4","new_1",
                                "new_14","new_9","new_0","new_2",
                                "null","null","new_16","new_15",
                                "new_fgby","new_12","new_7",
                                "new_8","new_11","new_10"};

        internal void doModel(List<ICell> list)
        {
            deptList = Lookup.RetrieveEntityAllRecord("new_dept", new String[] { });
            storeInformationList = Lookup.RetrieveEntityAllRecord("new_storeinformation", new String[] { "new_name" });

            new_6 = OptionSet.GetoptionsetText("new_dept", "new_6");
            new_5 = OptionSet.GetoptionsetText("new_dept", "new_5");
            new_17 = OptionSet.GetoptionsetText("new_dept", "new_17");
            new_13 = OptionSet.GetoptionsetText("new_dept", "new_13");
            new_4 = OptionSet.GetoptionsetText("new_dept", "new_4");
            new_14 = OptionSet.GetoptionsetText("new_dept", "new_14");
            new_9 = OptionSet.GetoptionsetText("new_dept", "new_9");
            new_16 = OptionSet.GetoptionsetText("new_dept", "new_16");
            new_12 = OptionSet.GetoptionsetText("new_dept", "new_12");
            new_7 = OptionSet.GetoptionsetText("new_dept", "new_7");
            new_8 = OptionSet.GetoptionsetText("new_dept", "new_8");
            new_11 = OptionSet.GetoptionsetText("new_dept", "new_11");
            new_10 = OptionSet.GetoptionsetText("new_dept", "new_10");
        }
        internal void deleteAll()
        {
            try
            {
                foreach (var d in deptList)
                {
                    service.Delete("new_dept", d.Id);
                }
            }
            catch (Exception ex)
            {
                EnvironmentSetting.ErrorMsg += "刪除資料失敗\n";
                EnvironmentSetting.ErrorMsg += ex.Message + "\n";
                EnvironmentSetting.ErrorMsg += ex.Source + "\n";
                EnvironmentSetting.ErrorMsg += ex.StackTrace + "\n";
                EnvironmentSetting.ErrorType = ErrorType.DATASYNCDETAIL;
            }
        }
        public TransactionStatus CreateForCRM(IRow row)
        {
            return AddAttributeForRecord(row, Guid.Empty);
        }
        private TransactionStatus AddAttributeForRecord(IRow row, Guid entityId)
        {
            try
            {
                Entity entity = new Entity("new_dept");

                ICell cell;
                Guid recordGuid;
                Int32 optionValue;
                for (int i = 0; i < fieldArray.Length; i++)
                {
                    cell = row.GetCell(i);
                    if (cell == null)
                        entity[fieldArray[i]] = null;
                    else
                    {
                        if (i == 0)
                        {
                            /// CRM欄位名稱     門市資訊    new_stonenumber
                            /// CRM關聯實體     門市資訊    new_storeinformation
                            /// CRM關聯欄位     店編        new_name
                            /// 
                            var test = from e in storeInformationList where e["new_name"].ToString() == cell.NumericCellValue.ToString() select e.Id;
                            recordGuid = test.First();
                            if (recordGuid == Guid.Empty)
                            {
                                EnvironmentSetting.ErrorMsg = "CRM 查無相符合資料 : \n";
                                EnvironmentSetting.ErrorMsg += "\tCRM實體 : new_storeinformation\n";
                                Console.WriteLine(EnvironmentSetting.ErrorMsg);
                                return TransactionStatus.Fail;
                            }
                            entity[fieldArray[i]] = new EntityReference("new_storeinformation", recordGuid);
                        }
                        else if (i == 2 || i == 3 || i == 16 || i == 17)
                        {

                        }
                        else if (i == 6)
                        {
                            if (cell.NumericCellValue == 0)
                                entity[fieldArray[i]] = null;
                            else
                                entity[fieldArray[i]] = cell.DateCellValue;
                        }
                        else
                        {
                            String value = "";
                            if (cell.CellType == CellType.Numeric)
                            {
                                value = cell.NumericCellValue.ToString();
                            }
                            else
                            {
                                value = cell.StringCellValue;
                            }

                            if (i == 4)
                            {
                                optionValue = OptionSet.getOptionSetValue(new_6, value);
                                if (optionValue == -1)
                                    entity[fieldArray[i]] = null;
                                else
                                    entity[fieldArray[i]] = new OptionSetValue(optionValue);
                            }
                            else if (i == 5)
                            {
                                optionValue = OptionSet.getOptionSetValue(new_5, value);
                                if (optionValue == -1)
                                    entity[fieldArray[i]] = null;
                                else
                                    entity[fieldArray[i]] = new OptionSetValue(optionValue);
                            }
                            else if (i == 8)
                            {
                                optionValue = OptionSet.getOptionSetValue(new_17, value);
                                if (optionValue == -1)
                                    entity[fieldArray[i]] = null;
                                else
                                    entity[fieldArray[i]] = new OptionSetValue(optionValue);
                            }
                            else if (i == 9)
                            {
                                optionValue = OptionSet.getOptionSetValue(new_13, value);
                                if (optionValue == -1)
                                    entity[fieldArray[i]] = null;
                                else
                                    entity[fieldArray[i]] = new OptionSetValue(optionValue);
                            }
                            else if (i == 10)
                            {
                                optionValue = OptionSet.getOptionSetValue(new_4, value);
                                if (optionValue == -1)
                                    entity[fieldArray[i]] = null;
                                else
                                    entity[fieldArray[i]] = new OptionSetValue(optionValue);
                            }
                            else if (i == 12)
                            {
                                optionValue = OptionSet.getOptionSetValue(new_14, value);
                                if (optionValue == -1)
                                    entity[fieldArray[i]] = null;
                                else
                                    entity[fieldArray[i]] = new OptionSetValue(optionValue);
                            }
                            else if (i == 13)
                            {
                                optionValue = OptionSet.getOptionSetValue(new_9, value);
                                if (optionValue == -1)
                                    entity[fieldArray[i]] = null;
                                else
                                    entity[fieldArray[i]] = new OptionSetValue(optionValue);
                            }
                            else if (i == 18)
                            {
                                optionValue = OptionSet.getOptionSetValue(new_16, value);
                                if (optionValue == -1)
                                    entity[fieldArray[i]] = null;
                                else
                                    entity[fieldArray[i]] = new OptionSetValue(optionValue);
                            }
                            else if (i == 21)
                            {
                                optionValue = OptionSet.getOptionSetValue(new_12, value);
                                if (optionValue == -1)
                                    entity[fieldArray[i]] = null;
                                else
                                    entity[fieldArray[i]] = new OptionSetValue(optionValue);
                            }
                            else if (i == 22)
                            {
                                optionValue = OptionSet.getOptionSetValue(new_7, value);
                                if (optionValue == -1)
                                    entity[fieldArray[i]] = null;
                                else
                                    entity[fieldArray[i]] = new OptionSetValue(optionValue);
                            }
                            else if (i == 23)
                            {
                                optionValue = OptionSet.getOptionSetValue(new_8, value);
                                if (optionValue == -1)
                                    entity[fieldArray[i]] = null;
                                else
                                    entity[fieldArray[i]] = new OptionSetValue(optionValue);
                            }
                            else if (i == 24)
                            {
                                optionValue = OptionSet.getOptionSetValue(new_11, value);
                                if (optionValue == -1)
                                    entity[fieldArray[i]] = null;
                                else
                                    entity[fieldArray[i]] = new OptionSetValue(optionValue);
                            }
                            else if (i == 25)
                            {
                                optionValue = OptionSet.getOptionSetValue(new_10, value);
                                if (optionValue == -1)
                                    entity[fieldArray[i]] = null;
                                else
                                    entity[fieldArray[i]] = new OptionSetValue(optionValue);
                            }
                            else
                            {
                                if (i == 7 || i == 11 || i == 14 || i == 15 || i == 19)
                                {
                                    entity[fieldArray[i]] = Convert.ToInt32(cell.NumericCellValue);
                                }
                                else
                                {
                                    if (value == "")
                                        entity[fieldArray[i]] = null;
                                    else
                                        entity[fieldArray[i]] = value;
                                }
                            }
                        }
                    }
                }
                try
                {
                    if (entityId == Guid.Empty)
                        service.Create(entity);
                    else
                    {
                        entity["new_deptid"] = entityId;
                        service.Update(entity);
                    }
                    return TransactionStatus.Success;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("店編 : " + row.GetCell(0).NumericCellValue.ToString());
                    Console.WriteLine(ex.Message);
                    EnvironmentSetting.ErrorMsg = ex.Message;
                    return TransactionStatus.Fail;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("欄位讀取錯誤");
                Console.WriteLine(ex.Message);
                EnvironmentSetting.ErrorMsg = "欄位讀取錯誤\n" + ex.Message;
                return TransactionStatus.Fail;
            }
        }

    }
}
