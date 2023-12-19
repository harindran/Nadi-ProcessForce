using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;
namespace Common.Common
{
    class clsTable
    {        
        public void FieldCreation()
        {
            
        }

        #region Table Creation Common Functions

        private void AddTables(string strTab, string strDesc, SAPbobsCOM.BoUTBTableType nType)
        {
            // var oUserTablesMD = default(SAPbobsCOM.UserTablesMD);
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;
            try
            {
                oUserTablesMD = (SAPbobsCOM.UserTablesMD)clsModule. objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                // Adding Table
                if (!oUserTablesMD.GetByKey(strTab))
                {
                    oUserTablesMD.TableName = strTab;
                    oUserTablesMD.TableDescription = strDesc;
                    oUserTablesMD.TableType = nType;

                    if (oUserTablesMD.Add() != 0)
                    {
                        throw new Exception(clsModule.objaddon.objcompany.GetLastErrorDescription() + strTab);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                oUserTablesMD = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private void AddFields(string strTab, string strCol, string strDesc, SAPbobsCOM.BoFieldTypes nType, int nEditSize = 10, SAPbobsCOM.BoFldSubTypes nSubType = 0, SAPbobsCOM.BoYesNoEnum Mandatory = SAPbobsCOM.BoYesNoEnum.tNO, string defaultvalue = "", bool Yesno = false, string[] Validvalues = null)
        {
            SAPbobsCOM.UserFieldsMD oUserFieldMD1;
            oUserFieldMD1 = (SAPbobsCOM.UserFieldsMD)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
            try
            {
             
                if (!IsColumnExists(strTab, strCol))
                {                   
                    oUserFieldMD1.Description = strDesc;
                    oUserFieldMD1.Name = strCol;
                    oUserFieldMD1.Type = nType;
                    oUserFieldMD1.SubType = nSubType;
                    oUserFieldMD1.TableName = strTab;
                    oUserFieldMD1.EditSize = nEditSize;
                    oUserFieldMD1.Mandatory = Mandatory;
                    oUserFieldMD1.DefaultValue = defaultvalue;

                    if (Yesno == true)
                    {
                        oUserFieldMD1.ValidValues.Value = "Y";
                        oUserFieldMD1.ValidValues.Description = "Yes";
                        oUserFieldMD1.ValidValues.Add();
                        oUserFieldMD1.ValidValues.Value = "N";
                        oUserFieldMD1.ValidValues.Description = "No";
                        oUserFieldMD1.ValidValues.Add();
                    }
                    //if (LinkedSystemObject != 0)
                    //    oUserFieldMD1.LinkedSystemObject = LinkedSystemObject;

                    string[] split_char;
                    if (Validvalues !=null)
            {
                        if (Validvalues.Length > 0)
                        {
                            for (int i = 0, loopTo = Validvalues.Length - 1; i <= loopTo; i++)
                            {
                                if (string.IsNullOrEmpty(Validvalues[i]))
                                    continue;
                                split_char = Validvalues[i].Split(Convert.ToChar(","));
                                if (split_char.Length != 2)
                                    continue;
                                oUserFieldMD1.ValidValues.Value = split_char[0];
                                oUserFieldMD1.ValidValues.Description = split_char[1];
                                oUserFieldMD1.ValidValues.Add();
                            }
                        }
                    }
                    int val;
                    val = oUserFieldMD1.Add();
                    if (val != 0)
                    {
                        clsModule.objaddon.objapplication.SetStatusBarMessage(clsModule. objaddon.objcompany.GetLastErrorDescription() + " " + strTab + " " + strCol, SAPbouiCOM.BoMessageTime.bmt_Short,true);
                    }
                    // System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1);
                oUserFieldMD1 = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private bool IsColumnExists(string Table, string Column)
        {
            SAPbobsCOM.Recordset oRecordSet=null;
            string strSQL;
            try
            {
               
                strSQL = "SELECT COUNT(*) FROM CUFD WHERE \"TableID\" = '" + Table + "' AND \"AliasID\" = '" + Column + "'";
                             
                oRecordSet = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordSet.DoQuery(strSQL);

                if (Convert.ToInt32( oRecordSet.Fields.Item(0).Value) == 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private void AddKey(string strTab, string strColumn, string strKey, int i)
        {
            var oUserKeysMD = default(SAPbobsCOM.UserKeysMD);

            try
            {
                // // The meta-data object must be initialized with a
                // // regular UserKeys object
                oUserKeysMD =(SAPbobsCOM.UserKeysMD)clsModule. objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys);

                if (!oUserKeysMD.GetByKey("@" + strTab, i))
                {

                    // // Set the table name and the key name
                    oUserKeysMD.TableName = strTab;
                    oUserKeysMD.KeyName = strKey;

                    // // Set the column's alias
                    oUserKeysMD.Elements.ColumnAlias = strColumn;
                    oUserKeysMD.Elements.Add();
                    oUserKeysMD.Elements.ColumnAlias = "RentFac";

                    // // Determine whether the key is unique or not
                    oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES;

                    // // Add the key
                    if (oUserKeysMD.Add() != 0)
                    {
                        throw new Exception(clsModule. objaddon.objcompany.GetLastErrorDescription());
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD);
                oUserKeysMD = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void AddUDO(string strUDO, string strUDODesc, SAPbobsCOM.BoUDOObjType nObjectType, string strTable, string[] childTable, string[] sFind, bool canlog = false, bool Manageseries = false)
        {

           SAPbobsCOM.UserObjectsMD oUserObjectMD=null;
            int tablecount = 0;
            try
            {
                oUserObjectMD = (SAPbobsCOM.UserObjectsMD)clsModule. objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
               
                if (!oUserObjectMD.GetByKey(strUDO)) //(oUserObjectMD.GetByKey(strUDO) == 0)
                {
                    oUserObjectMD.Code = strUDO;
                    oUserObjectMD.Name = strUDODesc;
                    oUserObjectMD.ObjectType = nObjectType;
                    oUserObjectMD.TableName = strTable;

                    oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES;

                    if (Manageseries)
                        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
                    else
                        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;

                    if (canlog)
                    {
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES;
                        oUserObjectMD.LogTableName = "A" + strTable.ToString();
                    }
                    else
                    {
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO;
                        oUserObjectMD.LogTableName = "";
                    }

                    oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.ExtensionName = "";

                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                    tablecount = 1;
                    if (sFind.Length > 0)
                    {
                        for (int i = 0, loopTo = sFind.Length - 1; i <= loopTo; i++)
                        {
                            if (string.IsNullOrEmpty(sFind[i]))
                                continue;
                            oUserObjectMD.FindColumns.ColumnAlias = sFind[i];
                            oUserObjectMD.FindColumns.Add();
                            oUserObjectMD.FindColumns.SetCurrentLine(tablecount);
                            tablecount = tablecount + 1;
                        }
                    }

                    tablecount = 0;
                    if (childTable != null)
            {
                        if (childTable.Length > 0)
                        {
                            for (int i = 0, loopTo1 = childTable.Length - 1; i <= loopTo1; i++)
                            {
                                if (string.IsNullOrEmpty(childTable[i]))
                                    continue;
                                oUserObjectMD.ChildTables.SetCurrentLine(tablecount);
                                oUserObjectMD.ChildTables.TableName = childTable[i];
                                oUserObjectMD.ChildTables.Add();
                                tablecount = tablecount + 1;
                            }
                        }
                    }

                    if (oUserObjectMD.Add() != 0)
                    {
                        throw new Exception(clsModule. objaddon.objcompany.GetLastErrorDescription());
                    }
                }
            }

            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                oUserObjectMD = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

        }


        #endregion


        

    }
}
