using CRM.Common;
using CRM.Model;
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
            deptList = Lookup.RetrieveEntityAllRecord("new_dept", new String[] { "new_stonenumber" });
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
        internal void deleteSameStonenumber(IRow row)
        {
            try
            {
                var value = row.GetCell(0).CellType == CellType.Numeric ? row.GetCell(0).NumericCellValue.ToString() : row.GetCell(0).StringCellValue;
                var smaeStronenumberList = (from s in deptList where ((EntityReference)s["new_stonenumber"]).Name.ToString() == value select s).ToList();
                foreach (var s in smaeStronenumberList)
                {
                    service.Delete("new_dept", s.Id);
                }
                deptList = deptList.Except(smaeStronenumberList).ToList();
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
                            var value = row.GetCell(0).CellType == CellType.Numeric ? row.GetCell(0).NumericCellValue.ToString() : row.GetCell(0).StringCellValue;
                            var temp = from e in storeInformationList where e["new_name"].ToString() == value select e.Id;
                            recordGuid = temp.First();
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
        internal void CalculateDept(DataSyncModel dataSync)
        {
            dataSync.CreateDataSyncForCRM("套組 - 更新門市資訊");
            if (EnvironmentSetting.ErrorType == ErrorType.None)
            {
                deptList = Lookup.RetrieveEntityAllRecord("new_dept", new String[] { "new_stonenumber", "new_9" });
                storeInformationList = Lookup.RetrieveEntityAllRecord("new_storeinformation", new String[] { "new_name" });

                int success = 0;
                int fail = 0;
                int partially = 0;
                foreach (var s in storeInformationList)
                {
                    try
                    {
                        Int32 new_gergf = 0;
                        Int32 new_feert = 0;
                        Int32 new_gftyu = 0;
                        Int32 new_fewtbhj = 0;
                        Int32 new_fewtr = 0;
                        Int32 new_vbgyuvb = 0;
                        Int32 new_vbgyumk = 0;
                        Int32 new_fewtrwqw = 0;
                        Int32 new_gerqws = 0;
                        Int32 new_fsetyt = 0;
                        Int32 new_fwetet = 0;
                        Int32 new_fwewhj = 0;

                        var list = (from e in deptList where ((EntityReference)e["new_stonenumber"]).Id == s.Id select ((OptionSetValue)e["new_9"]).Value).ToList();
                        foreach (var l in list)
                        {
                            var optionLabel = OptionSet.getOptionSetLabel(new_9, l);
                            new_gergf++;
                            switch (optionLabel)
                            {
                                case "調理": new_feert++; break;
                                case "日用": new_gftyu++; break;
                                case "休閒": new_fewtbhj++; break;
                                case "雜誌架": new_fewtr++; break;
                                case "資訊架": new_vbgyuvb++; break;
                                case "報架": new_vbgyumk++; break;
                                case "菸箱數量": new_fewtrwqw++; break;
                                case "推薦品架": new_gerqws++; break;
                                case "櫃前網架": new_fsetyt++; break;
                                case "黑色落地架": new_fwetet++; break;
                                case "零嘴落地架": new_fwewhj++; break;
                                default: break;
                            }
                        }

                        Entity entity = new Entity("new_storeinformation");

                        entity["new_storeinformationid"] = s.Id;

                        entity["new_gergf"] = new_gergf;
                        entity["new_feert"] = new_feert;
                        entity["new_gftyu"] = new_gftyu;
                        entity["new_fewtbhj"] = new_fewtbhj;
                        entity["new_fewtr"] = new_fewtr;
                        entity["new_vbgyuvb"] = new_vbgyuvb;
                        entity["new_vbgyumk"] = new_vbgyumk;
                        entity["new_fewtrwqw"] = new_fewtrwqw;
                        entity["new_gerqws"] = new_gerqws;
                        entity["new_fsetyt"] = new_fsetyt;
                        entity["new_fwetet"] = new_fwetet;
                        entity["new_fwewhj"] = new_fwewhj;

                        service.Update(entity);
                        success++;
                    }
                    catch
                    {
                        dataSync.CreateDataSyncDetailForCRM(s["new_name"].ToString(), s["new_name"].ToString(), TransactionType.Update, TransactionStatus.Fail);
                        fail++;

                        //新增detail錯誤 則結束
                        if (EnvironmentSetting.ErrorType != ErrorType.None)
                        {
                            return;
                        }
                    }
                }
            }
        }

    }
}
