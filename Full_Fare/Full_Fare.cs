using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Full_Fare.Model;
using System.IO;
using Excel;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Data.Entity.Validation;

namespace Full_Fare
{

  
    public partial class Full_Fare : Form
    {

        OpenFileDialog dialog;
        String strFileName, filePath;
        int rowCnt, columnCnt;
        DataTable table0;
        List<FullFare> fullFares_entity;

        public Full_Fare()
        {
            InitializeComponent();


        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            dialog = new OpenFileDialog();
            dialog.Filter ="xls files (*.xls)|*.xls|All files (*.*)|*.*";
            dialog.InitialDirectory = @"C:\jh\backup\CoreLogic";
            dialog.Title = "Select a xls file";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                strFileName = dialog.FileName;

                textBox1.Text = strFileName;
            }


            if (strFileName == String.Empty)
                return;
        }

        private void button2_Click(object sender, EventArgs e)
        {

            if (textBox1.Text == "") 
            {
              
                textBox2.AppendLine("No excel file specified, please choose a xls file try again");
              
                return;
            }

            if (File.Exists(textBox1.Text) == false)
	        {
               
                textBox2.AppendLine("File does not exist, please choose a xls file try again");
                return;
            }

            textBox2.Clear();
            
            ReadExcelToDataSet();
            
            fullFares_entity = MapDataSetToEntityModel();
            
            if (fullFares_entity != null)
                    { 
                        UpdateDB(fullFares_entity);
                    }


        }



        private void ReadExcelToDataSet()
        {

            filePath = textBox1.Text;

            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);


            IExcelDataReader excelReader = null;
            try
            {

                //1. Reading from a binary Excel file ('97-2003 format; *.xls)
                if (filePath.EndsWith(".xls"))
                {
                    textBox2.AppendLine("reading " + filePath);
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                }


                //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                if (filePath.EndsWith(".xlsx"))
                {
                    textBox2.AppendLine("reading " + filePath);
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }
                
            }
            catch (Exception excelreaderEX)
            {
                textBox2.AppendLine(excelreaderEX.Message);
                return;
            }

            //3. DataSet - The result of each spreadsheet will be created in the result.Tables
            //DataSet result = excelReader.AsDataSet();

            //4. DataSet - Create column names from first row
            excelReader.IsFirstRowAsColumnNames = true;
            DataSet result = excelReader.AsDataSet();


            if (result.Tables.Count > 1)
            {
                textBox2.AppendLine("Excel has multiple(" +  result.Tables.Count + ") worksheet please use one worksheet per file");
                return;
            }

            if (result.Tables.Count == 0 )
            {
                textBox2.AppendLine("Excel is empty, please use valid excel file");
                return;
            }

           table0 = result.Tables[0];

            rowCnt = table0.Rows.Count;
            columnCnt = table0.Columns.Count;

            textBox2.AppendLine("Excel Column (" + columnCnt + ")");
            textBox2.AppendLine("Excel Row (" + rowCnt + ")");


            // reformat each column name to comply with database table name / entity name
          
            string column0 = null;
            
            foreach (DataColumn column in table0.Columns)
            {
              
                column0 = column.ColumnName;
                
                if (column0.Contains(" "))
                {
                    column0 = Regex.Replace(column0, @"\s", "_");
                }

                if (column0.Contains("/"))
                {
                    column0 = Regex.Replace(column0, @"/", "_");
                }
                if (column0.Contains("+"))
                {
                    column0 = Regex.Replace(column0, @"\+", "_");
                }
                if (column0.Contains("."))
                {
                    column0 = Regex.Replace(column0, @"\.", "");
                }
                

                // handle special column name that do not match database table field name

                if (column0 == "FormattedApn") { column0 = "APN"; }

                if (column0 == "1st_Mtg_Amount") {column0 = "st_TD_Amount";}
                if (column0 == "1st_Mtg_Interest_Rate") { column0 = "st_TD_Int_Rate"; }
                if (column0 == "1st_Mtg_Interest_Rate_Type") { column0 = "st_TD_Int_Rate_Type"; }
                if (column0 == "1st_Mtg_Loan_Type") { column0 = "st_TD_Loan_Type"; }
                if (column0 == "1st_Mtg_Term_No") { column0 = "st_TD_Term_No"; }
                if (column0 == "Adjusted_Area_Total") { column0 = "Adj_Area__Total_"; }
                if (column0 == "Baths_Full") { column0 = "Baths__Full_"; }
                if (column0 == "Baths_Half") { column0 = "Baths__Half_"; }
                if (column0 == "Baths_Total") { column0 = "Baths__Total_"; }
                if (column0 == "Legal") { column0 = "Legal_Description"; }
                if (column0 == "Municipality_Township") { column0 = "Municipality"; }
                if (column0 == "Owner_Name_First_Name_First") { column0 = "Owner_Name__First_Name_First_"; }
                if (column0 == "Total_Value_Taxable") { column0 = "Total_Value__Taxable_"; }
                if (column0 == "Year_Built_Effective") { column0 = "Year_Built__Effective_"; }
                
                column.ColumnName = column0;
                table0.AcceptChanges();
               
            }


            

            //5. Data Reader methods
            //while (excelReader.Read())
            //{
            //    //excelReader.GetInt32(0);
            //}



            //6. Free resources (IExcelDataReader is IDisposable)
            excelReader.Close();

        }


        private List<FullFare> MapDataSetToEntityModel()
        {

            List<FullFare> fullFares = new List<FullFare>();
              //fullFares = new List<FullFare>();

            foreach (DataRow dtRow in table0.Rows)
            {

                var fullfare = new FullFare();
                foreach(DataColumn dtColumn in table0.Columns)
                {

                    var currentColumn = dtColumn.ColumnName;
                    Type t = fullfare.GetType();
                    PropertyInfo pf = t.GetProperty(currentColumn);
                    if (pf == null)
                    {
                        // property does not exist, entity property name does not match with datatable column name, need special handle.

                       // textBox2.AppendLine("Not Exist column : " + currentColumn);

                    }
                    else
                    {
                        // property exists

                        //var fieldValue = dtRow.Field<String>(currentColumn);
                        var fieldValue = dtRow[currentColumn];

                        
                        
                        var realType = pf.PropertyType;
                        // entity property with Nullable type, need to get underlying real type 
                        if (realType.IsGenericType && realType.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
                        {

                            realType = Nullable.GetUnderlyingType(realType);
                        }
                        
                        
                        // date type need to validate valid date, if not need to make up 1st day as the date.


                        if (realType == typeof(System.DateTime))
                        {
                            DateTime validDateTime;
                            String dateTimestring = fieldValue.ToString();
                            if (!DateTime.TryParse(dateTimestring, out validDateTime))
                            {
                                // invalid datatime string found
                                //textBox2.AppendLine("Invalid DateTime ---> " + dateTimestring);
                                
                                dateTimestring = dateTimestring.Replace("/00/", "/01/");
                                fieldValue = dateTimestring;


                                //textBox2.AppendLine("Corrected DateTime ---> " + dateTimestring);
                            }
                        }


                        // convert string to specific type according to entity property type
                        try
                        {
                            if ((string.IsNullOrEmpty(fieldValue.ToString())) || string.IsNullOrWhiteSpace(fieldValue.ToString()))
                            {
                                // can not set empty value into entity property. just skip setValue()
                            }
                            else
                            {
                                //pf.SetValue(fullfare, Convert.ChangeType(fieldValue, pf.PropertyType), null);
                                pf.SetValue(fullfare, WinFormsExtensions.ChangeType(fieldValue, pf.PropertyType), null);
                            }

                            }
                        catch (Exception typeConvertorSetValueError)
                        {

                            // display error message to log window either "convert type" or "set value" error out
                            textBox2.AppendLine("Value invalid :(" + fieldValue + ") should be " + realType + " *** error because of *** " + typeConvertorSetValueError);
                            textBox2.AppendLine(" Correct the excel file and try again");
                            return null;
                        }

                        //textBox2.AppendLine("Set " + currentColumn + " --> " + fieldValue);

                    } // else property exist.


                }// foreach dataColumn
                
                    
                // add one row to entities
                fullFares.Add(fullfare);

                

            } // foreach dataRow


            return fullFares;
        }


        //  Data access using database context
        private void UpdateDB(List<FullFare> fullFares_entities)
        {


           // database context
            using (var fullfareDBcontext = new FullFareEntities())
            {

                 // transaction with rollback 
                using (var dbcxtransaction = fullfareDBcontext.Database.BeginTransaction())
                {
                    try
                    {

                        foreach (FullFare listItem in fullFares_entities)
                        {
                            var apn = listItem.APN;

                            var query = from ff in fullfareDBcontext.FullFares
                                        where ff.APN.Equals(apn)
                                        select ff;


                            if (query.Count() > 0)
                            {
                                // apn exist
                                foreach (var existApn in query)
                                {
                                    // replace objectID with old record
                                    listItem.OBJECTID_1 = existApn.OBJECTID_1;


                                    textBox2.AppendLine("Existing APN record : " + existApn.APN + "  -- Address : " + existApn.ADDRESSPRP);
                                    fullfareDBcontext.FullFares.Remove(existApn);

                                }
                            }
                            else
                            {

                                //listItem.OBJECTID_1 = BitConverter.ToInt32(Guid.NewGuid().ToByteArray(), 0);
                                listItem.OBJECTID_1 = Math.Abs(Guid.NewGuid().GetHashCode());

                            }


                            fullfareDBcontext.FullFares.Add(listItem);
                            textBox2.AppendLine("Add: " + listItem.APN + "  -- Address : " + listItem.Situs_House_No_House_Alpha + " " + listItem.Situs_Direction_Street_Suffix);



                        }// foreach





                        fullfareDBcontext.SaveChanges();

                        dbcxtransaction.Commit();

                    }// try context transaction



                     // catch validation error only
                    catch (DbEntityValidationException e)
                    {

                       
                                // drill down the details of entity validation error for which item, which property cause validation error.
                                foreach (var eve in e.EntityValidationErrors)
                                {
                                    Console.WriteLine("Entity of type \"{0}\" in state \"{1}\" has the following validation errors:",
                                        eve.Entry.Entity.GetType().Name, eve.Entry.State);
                                    textBox2.AppendLine("Entity of type \"{0}\" in state \"{1}\" has the following validation errors: " + 
                                        eve.Entry.Entity.GetType().Name + " ---  " +  eve.Entry.State);

                                    foreach (var ve in eve.ValidationErrors)
                                    {
                                        Console.WriteLine("- Property: \"{0}\", Error: \"{1}\"", ve.PropertyName, ve.ErrorMessage);
                                        textBox2.AppendLine("- Property: \"{0}\" + Error: \"{1}\"" +  ve.PropertyName +  ve.ErrorMessage);
                                    }
                                }// foreach
                       

                        

                        //throw;
                    }// catch

                    // catch database insert error with rollback triggered
                    catch (Exception otherExp)
                    {

                        textBox2.AppendLine("Failed, rollback -- " + otherExp.ToString());
                        dbcxtransaction.Rollback();

                    }


                    textBox2.AppendLine("Update full fare database successfully ");

                }// using transaction

            }// using context




        }

    }






    public static class WinFormsExtensions
    {
        public static void AppendLine(this TextBox source, string value)
        {
            if (source.Text.Length == 0)
                source.Text = value;
            else
                source.AppendText("\r\n" + value);
        }



        




        public static object ChangeType(object value, Type conversion)
        {
            var t = conversion;

            // can not convert a null/ DBNull (value) to a generic type,
            if ((value == null) || (value is DBNull))
            {
                // no convert needed, just return null instead.
                return null;
            }

            else
            {
                if (t.IsGenericType && t.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
                {

                    t = Nullable.GetUnderlyingType(t);
                }


                return Convert.ChangeType(value, t);

            }//else
        }




        








    }



}
