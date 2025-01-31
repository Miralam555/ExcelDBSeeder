using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DimDb
{
    public class ExcelRead
    {

        public void ReadandWrite()
        {
            SqlCommand command = ConnectionDB.connect.CreateCommand();
            string excelFilePath = @"PATH";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string columnName = "COLUMNNAME";


            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {

                var worksheet = package.Workbook.Worksheets[0];


                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                int ColumnIndex = 0;
                int StartDateColumn = 8;
                int EndDateColumn = 11;
                string TaxID = null;
                int forTaxIdUnique = 1;

                for (int col = 1; col <= colCount; col++)
                {
                    if (worksheet.Cells[1, col].Text == columnName)
                    {
                        ColumnIndex = col;

                        for (int row = 2; row <= rowCount; row++)
                        {
                            var cellValue = worksheet.Cells[row, ColumnIndex].Value;
                            var cellMuessise = worksheet.Cells[row, ColumnIndex + 1].Value;
                            var cellMuqavileNo = worksheet.Cells[row, ColumnIndex - 1].Value;
                            var cellStartDate = worksheet.Cells[row, StartDateColumn].Value;
                            var cellEndDate = worksheet.Cells[row, EndDateColumn].Value;
                            var cellCategory = worksheet.Cells[row, 9].Value;
                            var AmountType = worksheet.Cells[row, 6].Value;
                            var cellCategoryType = worksheet.Cells[row, 10].Value;
                            var cellRecord = worksheet.Cells[row, 12].Value;
                            var cellSubject = worksheet.Cells[row, 4].Value;
                            var cellAmount = worksheet.Cells[row, 5].Value;
                            var cellPaymentMethod = worksheet.Cells[row, 13].Value;
                            TaxID = cellValue?.ToString();
                            bool TaxidHave = false;


                            if (TaxID != null)
                            {
                                command.CommandText = $"select * from Organizations where TaxId='{TaxID}'";
                                command.ExecuteNonQuery();

                                using (SqlDataReader read = command.ExecuteReader())
                                {
                                    if (read.Read())
                                    {
                                        TaxidHave = true;
                                    }
                                }
                            }



                            if (!TaxidHave)
                            {
                                command.CommandText = "insert into Organizations ([TaxID],[OrganizationName]) values (@TaxID,@OrganizationName)";

                                if (cellValue != null)
                                {
                                    if (TaxID.ToLower().Trim() == "yoxdur")
                                    {
                                        TaxID += (forTaxIdUnique++).ToString();//Taxidni unique etdim ve bunu oturecem bazaya
                                        command.Parameters.AddWithValue("@TaxID", TaxID);
                                    }
                                    else
                                    {
                                        command.Parameters.AddWithValue("@TaxID", cellValue.ToString());
                                    }


                                }
                                else
                                {
                                    command.Parameters.AddWithValue("@TaxID", DBNull.Value);
                                }
                                command.Parameters.AddWithValue("@OrganizationName", (string)cellMuessise);
                                command.ExecuteNonQuery();
                            }

                            command.Parameters.Clear();
                            int OrganizationId = 0;
                            command.CommandText = $"Select top 1 OrganizationID from Organizations where TaxID='{TaxID}'order by OrganizationId ";
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    OrganizationId = reader.GetInt32(0);
                                }


                            }
                            command.Parameters.Clear();
                            int AmountTypeId = 0;
                            switch (AmountType?.ToString().ToLower())
                            {
                                case "azn":
                                    AmountTypeId = 1;
                                    break;
                                case "avro":
                                    AmountTypeId = 2;
                                    break;
                                case "funt-sterlinq":
                                    AmountTypeId = 4;
                                    break;
                                case "dollar":
                                    AmountTypeId = 3;
                                    break;

                            }

                            command.CommandText = "insert into Contracts (ContractNumber,[Subject],[Amount],[AmountTypeId],[OrganizationId],[StartDate],[EndDate],[CategoryId],[TypeId],[StatusID],[PaymentMethodId]) values (@ContractNumber,@Subject,@Amount,@AmountTypeId,@OrganizationId,@StartDate,@EndDate,@CategoryId,@TypeId,@StatusID,@PaymentMethodId)";
                            command.Parameters.AddWithValue("@ContractNumber", cellMuqavileNo ?? DBNull.Value);
                            command.Parameters.AddWithValue("@OrganizationId", OrganizationId);
                            command.Parameters.AddWithValue("@Subject", cellSubject ?? DBNull.Value);
                            if (cellPaymentMethod != null)
                            {
                                if (cellPaymentMethod.ToString().ToLower().Trim()[0] == 'h')
                                {
                                    command.Parameters.AddWithValue("@PaymentMethodId", 1);
                                }
                                else if (cellPaymentMethod.ToString().ToLower().Trim()[0] == 't')
                                {
                                    command.Parameters.AddWithValue("@PaymentMethodId", 2);
                                }
                            }
                            else
                            {
                                command.Parameters.AddWithValue("@PaymentMethodId", DBNull.Value);
                            }
                            bool cellAmountControl = true;
                            if (cellAmount != null)
                            {
                                for (int i = 0; i < cellAmount.ToString().Length; i++)
                                {
                                    if (cellAmount.ToString().ToLower()[i] >= 'a' && cellAmount.ToString().ToLower()[i] <= 'z')
                                    {
                                        cellAmountControl = false;
                                        break;
                                    }
                                }
                                if (cellAmountControl)
                                {
                                    command.Parameters.AddWithValue("@Amount", Convert.ToDecimal(cellAmount));
                                    command.Parameters.AddWithValue("@AmountTypeId", AmountTypeId);
                                }
                                else
                                {
                                    command.Parameters.AddWithValue("@Amount", DBNull.Value);
                                    command.Parameters.AddWithValue("@AmountTypeId", DBNull.Value);
                                }
                            }
                            else
                            {
                                command.Parameters.AddWithValue("@Amount", DBNull.Value);
                                command.Parameters.AddWithValue("@AmountTypeId", DBNull.Value);
                            }
                            if (cellStartDate != null)
                            {
                                command.Parameters.AddWithValue("@StartDate", (DateTime)cellStartDate);
                            }
                            else
                            {
                                command.Parameters.AddWithValue("@StartDate", DBNull.Value);
                            }

                            string EndDateControl = cellEndDate?.ToString().Trim().ToLower();
                            if (EndDateControl != "öhdəlik" && cellEndDate != null)
                            {
                                if (cellEndDate is string)
                                {
                                    command.Parameters.AddWithValue("@EndDate", DateTime.ParseExact(cellEndDate.ToString(), "dd.MM.yyyy", CultureInfo.InvariantCulture));
                                }
                                else
                                {
                                    command.Parameters.AddWithValue("@EndDate", (DateTime)cellEndDate);
                                }

                            }
                            else
                            {
                                command.Parameters.AddWithValue("@EndDate", DBNull.Value);
                            }
                            int CategoryID = 0;
                            switch (cellCategory.ToString().Trim().ToLower())
                            {
                                case "xərc":
                                    command.Parameters.AddWithValue("@CategoryId", 1);
                                    CategoryID = 1;
                                    break;
                                case "gəlir":
                                    command.Parameters.AddWithValue("@CategoryId", 2);
                                    CategoryID = 2;
                                    break;
                                case "gəlir-xərc":
                                    command.Parameters.AddWithValue("@CategoryId", 3);
                                    CategoryID = 3;
                                    break;

                                case "ödənişsiz":
                                    command.Parameters.AddWithValue("@CategoryId", 4);
                                    CategoryID = 4;
                                    break;

                            }

                            /*
                             * Aradaki bosluqlar ucun alinmis ehtiyat
                             */
                            string[] Category = cellCategoryType.ToString().Trim().ToLower().Split(' ');
                            string TrueCategory = string.Empty;
                            for (int i = 0; i < Category.Length; i++)
                            {
                                TrueCategory += Category[i];
                            }
                            switch (TrueCategory)
                            {
                                case "satınalma-katirovka":
                                    command.Parameters.AddWithValue("@TypeId", 1);
                                    break;
                                case "satınalma-kotirovka":
                                    command.Parameters.AddWithValue("@TypeId", 1);
                                    break;
                                case "satınalma-tender":
                                    command.Parameters.AddWithValue("@TypeId", 2);
                                    break;
                                case "(3.8)büdcəplanıxaricisatınalma-xidməti":
                                    command.Parameters.AddWithValue("@TypeId", 4);
                                    break;
                                case "(3.8)büdcəplanıxaricisatınalma-alqı-satqı":
                                    command.Parameters.AddWithValue("@TypeId", 5);
                                    break;
                                case "(3.2.1)-dövlətsatınalmaistisnası":
                                    command.Parameters.AddWithValue("@TypeId", 6);
                                    break;

                                case "subpodrat":
                                    command.Parameters.AddWithValue("@TypeId", 8);
                                    break;

                                case "vesaitinayrılması":
                                    command.Parameters.AddWithValue("@TypeId", 10);
                                    break;
                                case "vəsaitinayrılması":
                                    command.Parameters.AddWithValue("@TypeId", 10);
                                    break;

                                case "icarə":
                                    command.Parameters.AddWithValue("@TypeId", 12);
                                    break;
                                case "imtahan":
                                    command.Parameters.AddWithValue("@TypeId", 13);
                                    break;

                                case "ödənişsiz":
                                    command.Parameters.AddWithValue("@TypeId", 16);
                                    break;

                            }
                            if (TrueCategory == "təlim" && CategoryID == 1)
                            {
                                command.Parameters.AddWithValue("@TypeId", 18);
                            }
                            if (TrueCategory == "təlim" && CategoryID == 2)
                            {
                                command.Parameters.AddWithValue("@TypeId", 19);
                            }
                            if (TrueCategory == "xidməti" && CategoryID == 1)
                            {
                                command.Parameters.AddWithValue("@TypeId", 7);
                            }
                            if (TrueCategory == "xidməti" && CategoryID == 2)
                            {
                                command.Parameters.AddWithValue("@TypeId", 11);
                            }
                            if (TrueCategory == "xidməti" && CategoryID == 3)
                            {
                                command.Parameters.AddWithValue("@TypeId", 17);
                            }
                            if (TrueCategory == "xidməti" && CategoryID == 4)
                            {
                                command.Parameters.AddWithValue("@TypeId", 15);
                            }
                            if (TrueCategory == "satınalma-birmənbə" && CategoryID == 1)
                            {
                                command.Parameters.AddWithValue("@TypeId", 3);
                            }
                            if (TrueCategory == "satınalma-birmənbə" && CategoryID == 2)
                            {
                                command.Parameters.AddWithValue("@TypeId", 9);
                            }


                            if ((cellEndDate != null && cellStartDate != null) && cellEndDate.ToString().Trim().ToLower()[0] != 'ö')
                            {
                                if (Convert.ToDateTime(cellEndDate) > Convert.ToDateTime(cellStartDate))
                                {
                                    if (cellRecord != null)
                                    {
                                        if (cellRecord.ToString().ToLower().Contains("uzadılma") || cellRecord.ToString().ToLower().Contains("uzadilma"))
                                        {
                                            command.Parameters.AddWithValue("@StatusID", 3);
                                        }
                                        else
                                        {
                                            command.Parameters.AddWithValue("@StatusID", 1);
                                        }
                                    }
                                    else
                                    {
                                        command.Parameters.AddWithValue("@StatusID", 1);
                                    }

                                }


                            }
                            else
                            {
                                command.Parameters.AddWithValue("@StatusID", 2);

                            }

                            command.ExecuteNonQuery();
                            command.Parameters.Clear();




                            int ContractId = 0;
                            command.CommandText = "Select top 1 Id from Contracts order by Id desc";
                            using (SqlDataReader reader1 = command.ExecuteReader())
                            {
                                if (reader1.Read())
                                {
                                    ContractId = reader1.GetInt32(0);
                                }
                            }
                            command.CommandText = "INSERT INTO Records ([RecordText],[ContractId]) values (@RecordText,@ContractId)";
                            if (cellRecord != null)
                            {
                                command.Parameters.AddWithValue("@RecordText", cellRecord);
                            }
                            else
                            {
                                command.Parameters.AddWithValue("@RecordText", DBNull.Value);
                            }

                            command.Parameters.AddWithValue("@ContractId", ContractId);
                            command.ExecuteNonQuery();
                            command.Parameters.Clear();




                        }
                    }
                }
                


                if (ColumnIndex == 0)
                {
                    Console.WriteLine("Sutun Tapilmadi");
                    return;
                }




            }

        }

    }
}
