using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using Oracle.DataAccess;
using Oracle.DataAccess.Client;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;


namespace ExcelDomainsSheets
{

  
    public partial class DomainsImportExport : Form
    {
        public DomainsImportExport()
        {
            InitializeComponent();
        }

        private Oracle.DataAccess.Client.OracleConnection DBConnectionDTOP ()

        {
            var connection = ConfigurationManager.ConnectionStrings["ExcelDomainsSheets.Properties.Settings.DTOP"].ConnectionString;
            Oracle.DataAccess.Client.OracleConnection conn = new OracleConnection(connection);
            return conn;
        }
        private Oracle.DataAccess.Client.OracleConnection DBConnectionCUBO()
        {
            var connection = ConfigurationManager.ConnectionStrings["ExcelDomainsSheets.Properties.Settings.CUBO_DESENV"].ConnectionString;
            Oracle.DataAccess.Client.OracleConnection conn = new OracleConnection(connection);
            return conn;
        }

        private void SQLToCSV(string nameOwner, string query, string domainFileName, string pathDomain)
        {
            Oracle.DataAccess.Client.OracleConnection conn = DBConnectionDTOP();
            conn.Open();
            OracleCommand cmd = new OracleCommand(query, conn);
            OracleDataReader dr = cmd.ExecuteReader();

            if (dr.HasRows)
            {
                CreateExcelFile(nameOwner, domainFileName, pathDomain, dr);
            }
            dr.Close();
            dr.Dispose();
            conn.Close();
        }

        private string MakeDomainsQuery(string domainOwner, string domainTable)
        {
            string strSQL;

            //strSQL = "SELECT ";
            //strSQL = strSQL + "ALLTABELAS.OWNER,  ";
            //strSQL = strSQL + "  TABELA.TABLE_NAME, ";
            //strSQL = strSQL + "  NULL, ";
            //strSQL = strSQL + "  COLUNAS.COLUMN_NAME AS COLUMN_NAME, ";
            //strSQL = strSQL + "  COMM.COMMENTS DESCRIPTION_COL, ";
            //strSQL = strSQL + "  COLUNAS.DATA_TYPE AS COLUMN_TYPE, ";
            //strSQL = strSQL + "  NVL(DECODE (NVL(COLUNAS.DATA_PRECISION,0), 0, COLUNAS.CHAR_COL_DECL_LENGTH, COLUNAS.DATA_PRECISION),0)  AS COLUMN_SIZE, ";
            //strSQL = strSQL + "  COLUNAS.DATA_SCALE AS DECIMAL_SIZE, ";
            //strSQL = strSQL + "  NULL, ";
            //strSQL = strSQL + "  COLUNAS.NULLABLE AS IS_NULL, ";
            //strSQL = strSQL + "  NVL(RELACIONA.PRIMARY_KEY, 'N') PRIMARY_KEY, ";
            //strSQL = strSQL + "  NULL, ";
            //strSQL = strSQL + "  NULL, ";
            //strSQL = strSQL + "  NULL, ";
            //strSQL = strSQL + "  NVL(RELACIONA.FOREIGN_KEY, 'N') FOREIGN_KEY, ";
            //strSQL = strSQL + "  FK.TABLE2 TABLE_RELATED, ";
            //strSQL = strSQL + "  FK.COLUMN2 COLUMN_REALATED ";
            //strSQL = strSQL + "FROM  USER_TABLES TABELA, ";
            //strSQL = strSQL + "      ALL_TABLES ALLTABELAS, ";
            //strSQL = strSQL + "      USER_TAB_COLUMNS COLUNAS, ";
            //strSQL = strSQL + "      ALL_COL_COMMENTS COMM, ";
            //strSQL = strSQL + "      (SELECT     ac.table_name, ";
            //strSQL = strSQL + "                  column_name, ";
            //strSQL = strSQL + "                  position, ";
            //strSQL = strSQL + "                  ac.constraint_name, ";
            //strSQL = strSQL + "                  DECODE (constraint_type, 'P', 'Y', 'N') PRIMARY_KEY, ";
            //strSQL = strSQL + "                  DECODE (constraint_type, 'P', 'N', 'Y') FOREIGN_KEY, ";
            //strSQL = strSQL + "                  (SELECT     ac2.table_name ";
            //strSQL = strSQL + "                  FROM       all_constraints ac2 ";
            //strSQL = strSQL + "                  WHERE      AC2.CONSTRAINT_NAME = AC.R_CONSTRAINT_NAME) fK_to_table ";
            //strSQL = strSQL + "      FROM       all_cons_columns acc, ";
            //strSQL = strSQL + "      all_constraints ac ";
            //strSQL = strSQL + "      WHERE      acc.constraint_name = ac.constraint_name AND ";
            //strSQL = strSQL + "                 acc.owner = ac.owner AND ";
            //strSQL = strSQL + "                 acc.table_name = ac.table_name AND ";
            //strSQL = strSQL + "                 CONSTRAINT_TYPE IN ('P', 'R')) RELACIONA, ";
            //strSQL = strSQL + "                 (select ";
            //strSQL = strSQL + "                       a.table_name TABLE1, ";
            //strSQL = strSQL + "                       max(decode(c.position,1,c.column_name)) COLUMN1, ";
            //strSQL = strSQL + "                       b.table_name TABLE2, ";
            //strSQL = strSQL + "                       max(decode(d.position,1,d.column_name)) COLUMN2 ";
            //strSQL = strSQL + "                    from ";
            //strSQL = strSQL + "                      user_constraints  a, ";
            //strSQL = strSQL + "                     user_constraints  b, ";
            //strSQL = strSQL + "                      user_cons_columns c, ";
            //strSQL = strSQL + "                     user_cons_columns d ";
            //strSQL = strSQL + "                   where ";
            //strSQL = strSQL + "                      a.r_constraint_name=b.constraint_name and ";
            //strSQL = strSQL + "                      a.constraint_name=c.constraint_name and ";
            //strSQL = strSQL + "                      b.constraint_name=d.constraint_name and ";
            //strSQL = strSQL + "                      a.constraint_type='R' and ";
            //strSQL = strSQL + "                      a.table_name = '" + domainTable + "' ";
            //strSQL = strSQL + "                   group by a.table_name, b.table_name ";
            //strSQL = strSQL + "                   order by 1) FK ";
            //strSQL = strSQL + "WHERE           TABELA.TABLE_NAME = COLUNAS.TABLE_NAME AND ";
            //strSQL = strSQL + "                ALLTABELAS.TABLE_NAME = TABELA.TABLE_NAME AND ";
            //strSQL = strSQL + "                RELACIONA.TABLE_NAME(+) = COLUNAS.TABLE_NAME  AND ";
            //strSQL = strSQL + "                RELACIONA.COLUMN_NAME(+) = COLUNAS.COLUMN_NAME AND ";
            //strSQL = strSQL + "                FK.TABLE1(+) = COLUNAS.TABLE_NAME AND ";
            //strSQL = strSQL + "                FK.COLUMN1(+) = COLUNAS.COLUMN_NAME AND ";
            //strSQL = strSQL + "                COMM.TABLE_NAME(+) = COLUNAS.TABLE_NAME AND ";
            //strSQL = strSQL + "                COMM.COLUMN_NAME(+) = COLUNAS.COLUMN_NAME AND ";
            //strSQL = strSQL + "                ALLTABELAS.OWNER = '" + domainOwner + "' AND  ";
            //strSQL = strSQL + "                TABELA.TABLE_NAME = '" + domainTable + "' ";
            //strSQL = strSQL + "order by  TABELA.TABLE_NAME, COLUNAS.COLUMN_ID ";  

            strSQL = "SELECT  ";
            strSQL = strSQL + "        ALLTABELAS.OWNER,  ";
            strSQL = strSQL + "        TABELA.TABLE_NAME, ";
            strSQL = strSQL + "        NULL, ";
            strSQL = strSQL + "        COLUNAS.COLUMN_NAME AS COLUMN_NAME, ";
            strSQL = strSQL + "        COMM.COMMENTS DESCRIPTION_COL, ";
            strSQL = strSQL + "        COLUNAS.DATA_TYPE AS COLUMN_TYPE, ";
            strSQL = strSQL + "        DECODE (NVL(COLUNAS.DATA_PRECISION,0), 0, COLUNAS.CHAR_COL_DECL_LENGTH, COLUNAS.DATA_PRECISION)  AS COLUMN_SIZE, ";
            strSQL = strSQL + "        NVL(COLUNAS.DATA_SCALE, 0) AS DECIMAL_SIZE, ";
            strSQL = strSQL + "        NULL, ";
            strSQL = strSQL + "        NULL, ";
            strSQL = strSQL + "        COLUNAS.NULLABLE AS IS_NULL, ";
            strSQL = strSQL + "        NVL((SELECT     'Y' AS PRIMARY_KEY ";
            strSQL = strSQL + "             FROM        all_cons_columns acc, ";
            strSQL = strSQL + "                        all_constraints ac ";
            strSQL = strSQL + "             WHERE       acc.constraint_name = ac.constraint_name AND ";
            strSQL = strSQL + "                        acc.owner = ac.owner AND ";
            strSQL = strSQL + "                        acc.table_name = ac.table_name AND ";
            strSQL = strSQL + "                       acc.table_name = COLUNAS.TABLE_NAME AND ";
            strSQL = strSQL + "                        acc.column_name = COLUNAS.COLUMN_NAME AND ";
            strSQL = strSQL + "                        CONSTRAINT_TYPE = ('P')), 'N') AS PRIMARYKEY, ";
            strSQL = strSQL + "         NULL, ";
            strSQL = strSQL + "         NULL, ";
            strSQL = strSQL + "         NULL, ";
            strSQL = strSQL + "         NVL((SELECT DISTINCT 'Y' AS FOREIGN_KEY ";
            strSQL = strSQL + "             FROM    all_cons_columns acc, ";
            strSQL = strSQL + "                     all_constraints ac ";
            strSQL = strSQL + "             WHERE   acc.constraint_name = ac.constraint_name AND ";
            strSQL = strSQL + "                     acc.owner = ac.owner AND ";
            strSQL = strSQL + "                     acc.table_name = ac.table_name AND ";
            strSQL = strSQL + "                     acc.table_name = COLUNAS.TABLE_NAME AND ";
            strSQL = strSQL + "                     acc.column_name = COLUNAS.COLUMN_NAME AND ";
            strSQL = strSQL + "                     CONSTRAINT_TYPE = ('R')), 'N') AS FOREIGNKEY, ";
            strSQL = strSQL + "         FK.CONSTR CONSTRAINT_NAME, ";
            strSQL = strSQL + "         FK.TB2 TABLE_RELATED, ";
            strSQL = strSQL + "         FK.C2 COLUMN_REALATED, ";
            strSQL = strSQL + "         FK.POS POSITION ";
            strSQL = strSQL + "FROM     USER_TABLES TABELA, ";
            strSQL = strSQL + "         ALL_TABLES ALLTABELAS, ";
            strSQL = strSQL + "         USER_TAB_COLUMNS COLUNAS, ";
            strSQL = strSQL + "         ALL_COL_COMMENTS COMM, ";
            strSQL = strSQL + "         (select distinct a.constraint_name as CONSTR, ";
            strSQL = strSQL + "                b.owner as OWNER, ";
            strSQL = strSQL + "                d.table_name as TB1, ";
            strSQL = strSQL + "                d.column_name as C1, ";
            strSQL = strSQL + "                b.table_name as TB2, ";
            strSQL = strSQL + "                b.column_name as C2, ";
            strSQL = strSQL + "                d.position as POS ";
            strSQL = strSQL + "         from user_constraints a, ";
            strSQL = strSQL + "                user_cons_columns b, ";
            strSQL = strSQL + "                user_constraints c, ";
            strSQL = strSQL + "                user_cons_columns d ";
            strSQL = strSQL + "         where a.table_name = '" + domainTable + "' ";
            strSQL = strSQL + "         and a.r_constraint_name = b.constraint_name ";
            strSQL = strSQL + "         and a.constraint_type = 'R' ";
            strSQL = strSQL + "         and a.constraint_name = c.constraint_name ";
            strSQL = strSQL + "         and c.constraint_name = d.constraint_name ";
            strSQL = strSQL + "         and b.position = d.position ";
            strSQL = strSQL + "         order by 1,7) FK ";
            strSQL = strSQL + "WHERE     TABELA.TABLE_NAME = COLUNAS.TABLE_NAME AND ";
            strSQL = strSQL + "          ALLTABELAS.TABLE_NAME = TABELA.TABLE_NAME AND ";
            strSQL = strSQL + "          FK.TB1(+) = COLUNAS.TABLE_NAME AND ";
            strSQL = strSQL + "          FK.C1(+) = COLUNAS.COLUMN_NAME AND ";
            strSQL = strSQL + "          COMM.TABLE_NAME(+) = COLUNAS.TABLE_NAME AND ";
            strSQL = strSQL + "          COMM.COLUMN_NAME(+) = COLUNAS.COLUMN_NAME AND ";
            strSQL = strSQL + "          ALLTABELAS.OWNER = '" +domainOwner + "' AND ";
            strSQL = strSQL + "          TABELA.TABLE_NAME = '" + domainTable + "' ";
            strSQL = strSQL + "order by  TABELA.TABLE_NAME, COLUNAS.COLUMN_ID ";

            return strSQL;

        }


        private string VerifyPrimaryKeyQuery(string domainOwner, string tableDomain)
        {
            string strSQL;

            strSQL = "SELECT ";
            strSQL = strSQL + "        ac.constraint_name ";
            strSQL = strSQL + "FROM    all_cons_columns acc, ";
            strSQL = strSQL + "        all_constraints ac ";
            strSQL = strSQL + "WHERE ";
            strSQL = strSQL + "        acc.constraint_name = ac.constraint_name AND ";
            strSQL = strSQL + "        acc.table_name = ac.table_name AND ";
            strSQL = strSQL + "        acc.owner = ac.owner AND ";
            strSQL = strSQL + "        acc.owner = '" + domainOwner + "' AND ";
            strSQL = strSQL + "        acc.table_name = '" + tableDomain + "' AND ";
            strSQL = strSQL + "        ac.CONSTRAINT_TYPE = 'P' ";

            return strSQL;

        }

        private string TablesForOwnerQuery(string domainOwner)
        {
            string strSQL;

            //strSQL = "SELECT  ";
            //strSQL = strSQL + "      ALLTABELAS.OWNER, ";
            //strSQL = strSQL + "      TABELA.TABLE_NAME ";
            //strSQL = strSQL + "FROM  USER_TABLES TABELA, ";
            //strSQL = strSQL + "      ALL_TABLES ALLTABELAS ";
            //strSQL = strSQL + "WHERE ";
            //strSQL = strSQL + "      ALLTABELAS.TABLE_NAME = TABELA.TABLE_NAME AND ";
            //strSQL = strSQL + "          ALLTABELAS.OWNER = '" + domainOwner + "' AND ";
            //strSQL = strSQL + "          ALLTABELAS.TABLE_NAME = 'CONTRATO_EMPRESA' ";
            //strSQL = strSQL + "order by  TABELA.TABLE_NAME ";

            strSQL = " select 'TS' as OWNER, ";
            strSQL = strSQL + " nome_tabela as TABLE_NAME ,";
            strSQL = strSQL + " lote as LOTE ";
            strSQL = strSQL + " from controle_dominio ";
            strSQL = strSQL + " where lote in (" + txtLote.Text + ") ";
            //strSQL = strSQL + " and nome_tabela in ('CTM_NR_CONTAS') ";
            strSQL = strSQL + " order by lote, TABLE_NAME ";
            return strSQL;

        }

        private string  RelatedDomainQuery(string columnType, int columnSize, int columnPrecision)

        {
            string strSQL;

            strSQL = "SELECT  dom.ds_domain_type ";
            strSQL = strSQL + "FROM    suggestions_columns_types sug, ";
            strSQL = strSQL + "      columns_types col, ";
            strSQL = strSQL + "      domains_types dom ";
            strSQL = strSQL + " WHERE   dom.id_domain_type = sug.id_domain_type AND ";
            strSQL = strSQL + "      col.id_column_type = sug.id_column_type AND ";
            strSQL = strSQL + "      (" + columnSize + " >= NVL(sug.nr_initial_range,0)) AND ";
            strSQL = strSQL + "      (" + columnSize + " <= NVL(sug.nr_end_range,0)) AND ";
            strSQL = strSQL + "      col.ds_column_type = '" + columnType + "' AND ";
            strSQL = strSQL + "      fl_decimal = 'N'";

            return strSQL;

        }


        private string DistinctsValuesQuery(string domainOwner, string nameTable, string nameColumn)
        {
            string strSQL;

            strSQL = "SELECT DISTINCT ";
            strSQL = strSQL + " " + nameColumn + " ";
            strSQL = strSQL + " FROM  " + domainOwner + "." + nameTable + " ";
            strSQL = strSQL + " WHERE " + nameColumn + " IS NOT NULL ";
            strSQL = strSQL + " AND ROWNUM < 5000 ";
 
            return strSQL;

        }

        private string OwnersQuery()
        {
            string strSQL;

            strSQL = "select distinct owner from all_tables order by owner ";

            return strSQL;

        }

        private bool VerifyPrimaryKey(string domainOwner, string tableDomain)
        {
            string strSQL = VerifyPrimaryKeyQuery(domainOwner, tableDomain);
            Oracle.DataAccess.Client.OracleConnection conn = DBConnectionDTOP();
            conn.Open();
            try
            {
                OracleCommand cmd = new OracleCommand(strSQL, conn);
                OracleDataReader dr = cmd.ExecuteReader();

                if (dr.HasRows)
                {
                    return true;
                }

                return false;
            }
            finally
            { conn.Close(); }
        }

        private void CreateExcelFile(string nameOwner, string domainTable, string excelSheetsPath, OracleDataReader dr)
        {
            bool fHavePrimaryKey;
            bool columnBoolean;
            //try
            //{
            
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);


                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                // headers columns
                xlWorkSheet.Cells[1, 1] = "Owner";
                xlWorkSheet.Cells[1, 2] = "Table Name";
                xlWorkSheet.Cells[1, 3] = "Class Domain Name";
                xlWorkSheet.Cells[1, 4] = "Column Name";
                xlWorkSheet.Cells[1, 5] = "Column Description";
                xlWorkSheet.Cells[1, 6] = "Column Type";
                xlWorkSheet.Cells[1, 7] = "Column Size";
                xlWorkSheet.Cells[1, 8] = "Column Precision";
                xlWorkSheet.Cells[1, 9] = "Property Name";
                xlWorkSheet.Cells[1, 10] = "Domain Type";
                xlWorkSheet.Cells[1, 11] = "Is Null";
                xlWorkSheet.Cells[1, 12] = "Is Primary Key";
                xlWorkSheet.Cells[1, 13] = "Is Boolean";
                xlWorkSheet.Cells[1, 14] = "Enumerator Name";
                xlWorkSheet.Cells[1, 15] = "Enumerator Values";
                xlWorkSheet.Cells[1, 16] = "Is Foreign Key";
                xlWorkSheet.Cells[1, 17] = "FK Constraint Name";
                xlWorkSheet.Cells[1, 18] = "FK Table Related";
                xlWorkSheet.Cells[1, 19] = "FK Column Related";
                xlWorkSheet.Cells[1, 20] = "FK Position";
                xlWorkSheet.Cells[1, 21] = "Reference Class";
                
                //Verify Primary Key Table
                fHavePrimaryKey = VerifyPrimaryKey(nameOwner, domainTable);

                // Loop through the rows and output the data
                int row = 1;
                while (dr.Read())
                {
                    row = row + 1;
                    columnBoolean = false;
                    for (int col = 0; col < dr.FieldCount; col++)
                    {
                        if (col != 12)
                        {
                            string value = dr[col].ToString();
                            xlWorkSheet.Cells[row, col + 1] = value;

                            if (col == 3 && dr[col].ToString().Length > 2)
                            {
                                 
                                if (dr[col].ToString().Substring(0, 3) == "IND")
                                {
                                      xlWorkSheet.Cells[row, 13] = "No identified. Verify!";
                                //    if (fHavePrimaryKey)
                                //    {
                                //        if (ColumnIsBoolean(nameOwner, domainTable, dr[col].ToString()))
                                //        {
                                //            xlWorkSheet.Cells[row, 13] = "Y";
                                //            columnBoolean = true;
                                //        }
                                //        else
                                //        {
                                //            xlWorkSheet.Cells[row, 13] = "No identified. Verify!";

                                //        }
                                //    }
                                //    else
                                //        {
                                //            xlWorkSheet.Cells[row, 13] = "Table without PK. Verify!";

                                //        }
                                }
                             }
                        }
                        if (col == 10)
                        {
                            string domainName = "";
                            if (!columnBoolean)
                            {
                                if (int.Parse(dr[7].ToString()) == 0)
                                {
                                    domainName = SearchDomain(dr[5].ToString(), (dr[6].ToString().Length == 0) ? 0 :int.Parse(dr[6].ToString()), int.Parse(dr[7].ToString()));
                                }
                                else
                                {
                                    domainName = "decimal";
                                }
                            }
                            else
                            {
                                domainName = "boolean";
                            }
                            xlWorkSheet.Cells[row, 10] = domainName;

                        }
                    }
                    
                }

                xlWorkBook.SaveAs(excelSheetsPath + "\\" + domainTable + ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
                        Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                ObjectsReleasing(xlWorkSheet);
                ObjectsReleasing(xlWorkBook);
                ObjectsReleasing(xlApp);

      

           // }
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Erro : " + ex.Message);
            //}
        }

        private string SearchDomain(string columnType, int columnSize, int columnPrecision)
        {
            string strSQL = RelatedDomainQuery(columnType, columnSize, columnPrecision);
            Oracle.DataAccess.Client.OracleConnection conn = DBConnectionCUBO();
            conn.Open();
            try
            {
                OracleCommand cmd = new OracleCommand(strSQL, conn);
                OracleDataReader dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    return dr[0].ToString();
                }
                return "";
            }
            finally {conn.Close();}
            
        }

        private bool ColumnIsBoolean(string nameOwner, string nameTable, string nameColumn)
        {
                List<string> listIndicateColumns = new List<string>();

                string strSQL = DistinctsValuesQuery(nameOwner, nameTable, nameColumn);
                Oracle.DataAccess.Client.OracleConnection conn = DBConnectionDTOP();
   
                conn.Open();
                try
                {
                    OracleCommand cmd = new OracleCommand(strSQL, conn);
                    OracleDataReader dr = cmd.ExecuteReader();
                    while (dr.Read())
                    {

                        listIndicateColumns.Add(dr[0].ToString());
                    }
                    if (listIndicateColumns.Count <= 2 && (listIndicateColumns.Contains("S") || listIndicateColumns.Contains("N")))
                    {
                        return true;
                    }
                    return false;
                }
                finally { conn.Close(); }
        }

        private void ObjectsReleasing(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("An error occurred while releasing the object." + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void DomainsImportExport_Load(object sender, EventArgs e)
        {
            FillsOwnerList();
        }

        private void FillsOwnerList()
        {
            string strSQL = "";
            strSQL = OwnersQuery();

            Oracle.DataAccess.Client.OracleConnection conn = DBConnectionDTOP();
            conn.Open();
            OracleCommand cmd = new OracleCommand(strSQL, conn);
            OracleDataReader dr = cmd.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(dr);
            
            cboOwner.DisplayMember = "Owner";
            cboOwner.ValueMember = "Owner";
            cboOwner.DataSource = dt;

            conn.Close();
            
        }
        
        private void btnCreateSheets_Click(object sender, EventArgs e)
        {

            if (cboOwner.SelectedIndex == -1)
            {
                MessageBox.Show("Please, select a owner for continue.");
                return;
            }

            string domainPathSource = "C:\\ExcelSheets\\TEMP\\" + cboOwner.SelectedValue.ToString();
                
            string strSQL = "";
            int count = 0;
            strSQL = TablesForOwnerQuery(cboOwner.SelectedValue.ToString());

            Oracle.DataAccess.Client.OracleConnection conn = DBConnectionDTOP();
            conn.Open();
            try
            {
                OracleCommand cmd = new OracleCommand(strSQL, conn);
                OracleDataReader dr = cmd.ExecuteReader();

                DataTable dt = new DataTable();
                dt.Load(dr);
                int totalLines = dt.Rows.Count;
                dt.Dispose();

                progressBar2.Value = 0;     // Esse é o valor da progress bar ela vai de 0 a Maximum (padrão 100)
                progressBar2.Maximum = totalLines;

                dr.Close();

                dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    
                    string domainPathTarget = "C:\\ExcelSheets\\TEMP\\" + cboOwner.SelectedValue.ToString() + "\\" + dr[2].ToString();


                    if (!System.IO.Directory.Exists(domainPathTarget))
                    {
                        System.IO.Directory.CreateDirectory(domainPathTarget);
                    }

                    if (System.IO.File.Exists(domainPathSource + "\\" + dr[1].ToString() + ".xls"))
                    {
                        System.IO.File.Move(domainPathSource + "\\" + dr[1].ToString() + ".xls", domainPathTarget + "\\" + dr[1].ToString() + ".xls");
                    }
                    else
                    {

                        label2.Text = "Exportando tabela: " + cboOwner.SelectedValue.ToString() + "." + dr[1].ToString() + " para planilha: " + domainPathTarget + "\\" + dr[1].ToString() + ".xls";
                        label2.Visible = true;

                        strSQL = MakeDomainsQuery(cboOwner.SelectedValue.ToString(), dr[1].ToString());
                        SQLToCSV(cboOwner.SelectedValue.ToString(), strSQL, dr[1].ToString(), domainPathTarget);

                    }
                    count++;
                    progressBar2.Value = count;
                    progressBar2.Visible = true;

                    label3.Text = "Exportada: " + count.ToString() + " de " + totalLines + " tabelas do Owner " + cboOwner.SelectedValue.ToString();
                    label3.Visible = true;

                }
                label3.Text = "Exportação realizada com sucesso para : " + count.ToString() + " tabelas do Owner " + cboOwner.SelectedValue.ToString();

            }
            finally { conn.Close(); conn.Dispose(); }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

    }
}
