using System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Windows.Forms;
using AddinExpress.MSO;
using System.Xml;
using Microsoft.Office.Interop.Excel;






namespace InfoselXLLAddin
{
    /// <summary>
    ///   Add-in Express XLL Add-in Module
    /// </summary>
    [ComVisible(true)]
    public partial class XLLModule : AddinExpress.MSO.ADXXLLModule
    {
        public XLLModule()
        {
           //Application.EnableVisualStyles(); RGC Comente esta linea.
            InitializeComponent();
            // Please add any initialization code to the OnInitialize event handler
        }
 
        #region Add-in Express automatic code
 
        // Required by Add-in Express - do not modify
        // the methods within this region
 
        public override System.ComponentModel.IContainer GetContainer()
        {
            if (components == null)
                components = new System.ComponentModel.Container();
            return components;
        }
 
        [ComRegisterFunctionAttribute]
        public static void RegisterXLL(Type t)
        {
            AddinExpress.MSO.ADXXLLModule.RegisterXLLInternal(t);
        }
 
        [ComUnregisterFunctionAttribute]
        public static void UnregisterXLL(Type t)
        {
            AddinExpress.MSO.ADXXLLModule.UnregisterXLLInternal(t);
        }
 
        #endregion
 
        public static new XLLModule CurrentInstance
        {
            get
            {
                return AddinExpress.MSO.ADXXLLModule.CurrentInstance as XLLModule;
            }
        }

        public Excel._Application ExcelApp
        {
            get
            {
                return (HostApplication as Excel._Application);
            }
        }

        #region Define your UDFs in this section
 
        /// <summary>
        /// The container for user-defined functions (UDFs). Every UDF is a public static (Public Shared in VB.NET) method that returns a value of any base type: string, double, integer.
        /// </summary>
        internal static class XLLContainer
        {
            /// <summary>
            /// Required by Add-in Express. Please do not modify this method.
            /// </summary>
            internal static XLLModule Module
            {
                get
                {
                    return AddinExpress.MSO.ADXXLLModule.
                        CurrentInstance as InfoselXLLAddin.XLLModule;
                }
            }
 
            #region Sample function
 

          
           
            public static int IQL_Suma(int a, int b)
            {
                return a + b;
            }

            public static int IQL_Resta(int a, int b)
            {
                return a - b;
            }

            public static int IQL_Multiplica(int a, int b)
            {
                return a * b;
            }

            public static int IQL_Divide(int a, int b)
            {
                return a / b;
            }

            public static int IQL_Cuadrado(int a)
            {
                return a * a;
            }

            // Seleccionar el area, despues la funcion, alimentar parametros y despues <CTRL><ALT><ENTER>
            // El area seleccionada debe de ser mayor a los parametros de entrada.
            public static object IQL_Arreglo(int rows, int columns)
            {
                object[,] result = new string[rows, columns];
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < columns; j++)
                    {
                        result[i, j] = string.Format("({0},{1})", i, j);
                    }
                }

                return result;
            }


            public static object IQL_ArregloMas(object[,] array, decimal numero)
            {
                int rows = array.GetLength(0);
                int columns = array.GetLength(1);

                object[,] result = new string[rows, columns];
              
                string straux = "";
                decimal dValor = 0;

                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < columns; j++)
                    {
                        straux = array[i, j].ToString();
                        decimal.TryParse(straux, out dValor);
                        dValor += numero;
                        result[i, j] = dValor.ToString();
                      }
                }

                return result;
            }


            public static object IQL_GetLast(string Simbolo)
            {
                String sURL = "http://finanzasenlinea.infosel.com/wsNeural/service.asmx/GetLastBySymbol?Symbol=" + Simbolo + "&Market=BMV&SymbolType=ACCION";
                object[,] result = new string[1, 8];

                XmlDocument xdoc = new XmlDocument();
                xdoc.Load(sURL);
                
                result[0, 0] = xdoc.DocumentElement.SelectNodes("/Quote")[0]["Emisora"].InnerText;
                result[0, 1] = xdoc.DocumentElement.SelectNodes("/Quote")[0]["UltimoHecho"].InnerText;
                result[0, 2] = xdoc.DocumentElement.SelectNodes("/Quote")[0]["ValorApertura"].InnerText;
                result[0, 3] = xdoc.DocumentElement.SelectNodes("/Quote")[0]["MaximoDia"].InnerText;
                result[0, 4] = xdoc.DocumentElement.SelectNodes("/Quote")[0]["MinimoDia"].InnerText;
                result[0, 5] = xdoc.DocumentElement.SelectNodes("/Quote")[0]["VolumenAcumulado"].InnerText;
                result[0, 6] = xdoc.DocumentElement.SelectNodes("/Quote")[0]["Fecha"].InnerText;
                result[0, 7] = xdoc.DocumentElement.SelectNodes("/Quote")[0]["Hora"].InnerText;

                return result;
            }



            public static object IQL_GetLast2(string Simbolo)
            {                
                return 0;            
            }            


            // http://stackoverflow.com/questions/17476620/returning-array-in-excel-xll








            #endregion
















            /*
                dgvResultado.Rows[1].Cells[1].Value = "New Valuev 1";
                dgvResultado.UpdateCellValue(1, 1);

                dgvResultado.Rows[1].Cells[2].Value = "New Valuev 2";
                dgvResultado.UpdateCellValue(1, 2);
*/

            /*
            DataGridViewRow row = (DataGridViewRow)dgvResultado.RowTemplate.Clone();
            row.CreateCells(dgvResultado, "HOLA","COMO", "ESTAS");
            dgvResultado.Rows.Add(row); 
            */

            //DataGridCell[,] x = new DataGridCell[1,2];
            //x[0, 0] = "a";
            //x[0, 1] = "a";

            //DataGridView dgv;
            //DataGridViewColumn[] dgvColumn = new DataGridViewColumn[2];
            //DataGridViewRow dgvRow;


            //int value;
            //if(int.TryParse(dgvRow.Cells[1].FormattedValue, out value))


            // dgvColumn[1].

            // Value = 100;
            // dgvColumn[2].Value = 200;

            //CurrentCell.Value = newValue;
            //DataGridCell x;
            //x. = "aaaaa";



            /*
                XmlNodeList xNodelst = xdoc.DocumentElement.SelectNodes("Quote");
                foreach (XmlNode xNode in xNodelst)
                {
                    sEmisora = xNode["Emisora"].InnerText;
                }
            */
                
  


        }
 
        #endregion
    }
}

