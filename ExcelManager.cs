using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using Tools = Microsoft.Office.Tools.Excel;

namespace DeltaCore.FileAccess {

  public class ExcelManager {
    #region Variables de Clase
		public Excel.Application ExApp;
		public Excel.Workbook ExWB;
		public Excel.Worksheet ExWS;
		public Excel.Sheets ExS;
		public Excel.Range ExRange;
  	
		protected object sinvalor = System.Reflection.Missing.Value;
  	public int SpaceRows { get; set; }
		public int SpaceColumns { get; set; }
    public int Row { get; set; }
		#endregion

		#region Inicial
		public ExcelManager() {
      try {
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("EN-US");
      	ExApp = new Excel.Application {DisplayAlerts = false, Visible = false};
				GC.Collect();
      }
      catch (COMException cex) {
				ExApp = null;
      }
      catch (Exception ex) {
        throw new Exception(ex.Message + "Error en Excel - No se puede crear el objeto Excel", ex);
      }
    }

		private void ThisWorkbook_Startup(object sender, System.EventArgs e)
		{
			//((Excel.AppEvents_Event)this.Application).NewWorkbook += new Excel.AppEvents_NewWorkbookEventHandler(ThisWorkbook_NewWorkbook);
		}

		private void ThisWorkbook_Shutdown(object sender, System.EventArgs e) { }

		// Inicia una instancia de una aplicación de Excel en blanco.	
		public Excel.Application AbrirArchivoExcel()
		{
			GC.Collect();
			ExWB = ExApp.Workbooks.Add(sinvalor);
			ExS = ExWB.Worksheets;
			ExWS = (Excel.Worksheet)ExApp.ActiveWorkbook.Worksheets[1];
			SpaceColumns = 5;
			SpaceRows = 5;
			return ExApp;
		}

		// Inicia una instancia de una aplicación de Excel           
		public Excel.Application AbrirArchivoExcel(string ruta)
		{
			ExWB = ExApp.Workbooks.Open(ruta, sinvalor, sinvalor, sinvalor, sinvalor, sinvalor, sinvalor, sinvalor, sinvalor, sinvalor, sinvalor, sinvalor, sinvalor, sinvalor, sinvalor);
			ExS = ExWB.Worksheets;
			ExWS = (Excel.Worksheet) ExApp.ActiveWorkbook.ActiveSheet;
			SpaceColumns = 5;
			SpaceRows = 5;
			return ExApp;
		}

		// Inicia una instancia de una aplicación de Excel permitiendo seleccionar la hoja activa.            
		public Excel.Application AbrirArchivoExcel(string ruta, string nombreHoja)
		{
			ExWB = ExApp.Workbooks.Open(ruta, sinvalor, sinvalor, sinvalor, sinvalor, sinvalor, sinvalor, sinvalor, sinvalor, sinvalor, sinvalor, sinvalor, sinvalor, sinvalor, sinvalor);
			ExS = ExWB.Worksheets;
			ExWS = (Excel.Worksheet)ExApp.ActiveWorkbook.Worksheets[nombreHoja];
			SpaceColumns = 5;
			SpaceRows = 5;
			return ExApp;
		}

    public void OpenTemplate(string rootPath, string template) {
      try {
        //ExLibro ExWB = ExApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet); Templete en blanco
        string workbookPath = rootPath + template;
        ExWB = ExApp.Workbooks.Open(workbookPath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
        //ExApp.Workbooks.Open(workbookPath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
				ExWS = (Excel.Worksheet)ExApp.ActiveSheet;
				SpaceColumns = 5;
				SpaceRows = 5;
      }
      catch (Exception ex) {
        throw new Exception(ex.Message + "Error en Excel - No se puede crear el objeto Excel", ex);
      }
    }

    public void OpenTemplate(string workbookPath, bool ReadOnly) {
      try {
        //ExLibro ExWB = ExApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet); Templete en blanco        
        ExApp.Workbooks.Open(workbookPath, 0, ReadOnly, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
				ExWS = (Excel.Worksheet)ExApp.ActiveSheet;
				SpaceColumns = 5;
				SpaceRows = 5;
      }
      catch (Exception ex) {
        throw new Exception(ex.Message + "Error en Excel - No se puede crear el objeto Excel", ex);
      }
    }

		public void Ready()
		{
			try
			{
				ExApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
				ExWS = (Excel.Worksheet)ExApp.ActiveSheet;
			}
			catch (Exception ex)
			{
				throw new Exception(ex.Message + "Error en Excel - No se puede crear el objeto Excel", ex);
			}
		}

    public void Presentacion(string Hoja, string Reporte, string Titulos) {
      try {
        //En el objeto Excel Application, se agrega un Workbook que es tomado por el objeto Excel Workbook        
        ExWB = ExApp.Workbooks.Add(Missing.Value);
        //Se añade un worksheet
        AgregarHoja(Hoja, Reporte, Titulos);
      }
      catch (COMException cex) {
        ExApp = null;
        ExWB = null;
      }
      catch (Exception ex) {
        throw new Exception(ex.Message + "Error en Excel - No se puede crear el objeto Excel", ex);
      }
    }

    public void Presentacion(string Hoja, XmlDocument xReporte) {
      try {
        //En el objeto Excel Application, se agrega un Workbook que es tomado por el objeto Excel Workbook        
        ExWB = ExApp.Workbooks.Add(Missing.Value);
        //Se añade un worksheet
        AgregarHoja(Hoja, xReporte);
      }
      catch (COMException cex) {
        ExApp = null;
        ExWB = null;
      }
      catch (Exception ex) {
        throw new Exception(ex.Message + "Error en Excel - No se puede crear el objeto Excel", ex);
      }
    }

    public void AgregarHoja(string Hoja) {
      Row = 1; //(int)Rows.RowTitle;
      try {
        if (ExWB == null)
          ExWB = ExApp.Workbooks.Add(Missing.Value);
        //Del objeto Excel Application, se agrega una hoja con parametros vacios
        ExApp.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
        //La hoja agregada, queda como hoja activa, para el objeto Excel Worksheet
				ExWS = (Excel.Worksheet)ExWB.ActiveSheet;
        //Se inserta el nombre a la hoja activa
        ExWS.Name = Hoja;
      }
      catch (Exception ex) {
        ExApp = null;
        ExWS = null;
        throw new Exception(ex.Message + "Error en Excel - No se puede crear el objeto Excel", ex);
      }
    }
    
    /// <summary>
    /// Este método Agrega una hoja del Excel Application al Excel Workbook.
    /// </summary>
    /// <param name="Hoja">El nombre que tendrá la Worksheet</param>
    /// <param name="xReporte">Es un XML con los encabezados que lleva el reporte, los tags determinan el despliegue</param>
    public void AgregarHoja(string Hoja, XmlDocument xReporte) {
			Row = 1; //(int)Rows.RowTitle;
      try {
        if (ExWB == null)
          ExWB = ExApp.Workbooks.Add(Missing.Value);
        //Del objeto Excel Application, se agrega una hoja con parametros vacios
        ExApp.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
        //La hoja agregada, queda como hoja activa, para el objeto Excel Worksheet
				ExWS = (Excel.Worksheet)ExWB.ActiveSheet;
        //Se inserta el nombre a la hoja activa
        ExWS.Name = Hoja;
        //Titulo del reporte
        ExTitulo(ExWS, xReporte);
      }
      catch (Exception ex) {
        ExApp = null;
        ExWS = null;
        throw new Exception(ex.Message + "Error en Excel - No se puede crear el objeto Excel", ex);
      }
    }

    /// <summary>
    /// Este método Agrega una hoja del Excel Application al Excel Workbook
    /// </summary>
    /// <param name="Hoja">El nombre que tendrá la Worksheet</param>
    /// <param name="Reporte">Titulo del Reporte que se despliega</param>
    /// <param name="Titulos">Son los titulos que lleva el reporte como encabezados</param>
    public void AgregarHoja(string Hoja, string Reporte, string Titulos) {
			Row = 1; //(int)Rows.RowTitle;
      try {
        if (ExWB == null)
          ExWB = ExApp.Workbooks.Add(Missing.Value);
        //Del objeto Excel Application, se agrega una hoja con parametros vacios
        ExApp.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
        //La hoja agregada, queda como hoja activa, para el objeto Excel Worksheet
				ExWS = (Excel.Worksheet)ExWB.ActiveSheet;
        //Se inserta el nombre a la hoja activa
        ExWS.Name = Hoja;
        //Titulo del reporte
        ExTitulo(ExWS, Reporte);
        //Insertar encabezados
        ExEncabezado(ExWS, 5, Titulos);
      }
      catch (COMException cex) {
        ExApp = null;
        ExWS = null;
      }
      catch (Exception ex) {
        throw new Exception(ex.Message + "Error en Excel - No se puede crear el objeto Excel", ex);
      }
    }

    /// <summary>
    /// Este método Agrega una hoja del Excel Application al Excel Workbook. Este método es ocupado por los exceles multihoja
    /// </summary>
    /// <param name="Hoja">El nombre que tendrá la Worksheet</param>
    /// <param name="Reporte">Es el encabezado principal del reporte</param>
    /// <param name="Comentario">Comentarios sobre el reporte que se depliega un renglon inmediato al reporte</param>
    /// <param name="Titulos">Son los encabezados del reporte en bloque</param>
    public void AgregarHoja(string Hoja, XmlNode Reporte, string Comentario, List<object> Titulos) {
			Row = 1; //(int)Rows.RowTitle;
      try {
        if (ExWB == null)
          ExWB = ExApp.Workbooks.Add(Missing.Value);
        //Del objeto Excel Application, se agrega una hoja con parametros vacios
        ExApp.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
        //La hoja agregada, queda como hoja activa, para el objeto Excel Worksheet
				ExWS = (Excel.Worksheet)ExWB.ActiveSheet;
        //Se inserta el nombre a la hoja activa
        ExWS.Name = Hoja;
        //Titulo del reporte
        ExTitulo(ExWS, Reporte, Comentario);
        //Insertar encabezados
        XmlElement xElement = (XmlElement)Reporte;
        ExEncabezado(ExWS, Titulos, xElement.Attributes["Interior"].Value);
      }
      catch (COMException cex) {
        ExApp = null;
        ExWS = null;
      }
      catch (Exception ex) {
        throw new Exception(ex.Message + "Error en Excel - No se puede crear el objeto Excel", ex);
      }
    }

    public string NombresHojas(string Archivo, byte X, bool Cerrar) {
      string Hojas = string.Empty;
      //Creo el objeto si es necesario, sino, toma el objeto ya creado    
      ExApp.Workbooks.Open(Archivo, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
      X = (byte)ExApp.Worksheets.Count;
      for (byte i = 1; i <= X; i++) {
        //Hojas = Hojas + ExApp.Worksheets[I].Name + ',';  
				var ws = new Excel.Worksheet();
        ws = (Excel.Worksheet)ExWB.Worksheets[i];
        Hojas += ws.Name + ',';
      }
      if (Cerrar)
        Finalizar();
      return Hojas;
    }
    #endregion

    #region Reader
    public int maxColumn(int Row) {
      return LastData(Row, false, SpaceColumns);
    }

    public int maxRow(int Column) {      
      return LastData(Column, true, SpaceRows);
    }

    public int maxRow(int Row, int Column) {      
      string cell = Ref2Cel(Row, Column);
      ExWS.Range[cell, cell].Activate();
      var ExR = (Excel.Range)ExApp.Selection;
      var ExRange = ExR.End[Excel.XlDirection.xlToLeft];
      int i = ExRange.Cells.Count;      
      ExRange = ExR.End[Excel.XlDirection.xlDown];
      return ExRange.Cells.Count;
    }

    public string LeerRenglon(int Renglon, bool Cerrar) {
      string Encabezados = string.Empty;
      string R1, S;
      int Y;
      int X = 10;
      S = ExApp.ActiveWorkbook.Name;
      Y = Renglon == 0 ? 1 : Renglon;
      R1 = Ref2Cel((byte)X, Y);
      while (ExWS.Range[R1, R1].Text.ToString() != string.Empty) {
        X = X + SpaceRows;
        R1 = Ref2Cel(X, Y);
      }
      while (ExWS.Range[R1, R1].Text.ToString() != string.Empty) {
        X--;
        R1 = Ref2Cel(X, Y);
      }
      for (int i = 1; i <= X; i++) {
        R1 = Ref2Cel(i, Y);
        Encabezados += Encabezados.Trim() + ',' + ExWS.Range[R1, R1].Text.ToString() + ',';
      }
      if (Cerrar)
        Finalizar();
      return Encabezados;
    }

    public object ReaderCell(string celda) {
      return ExWS.Range[celda, celda].Text;
    }

    public object ReaderCell(int celda1, int celda2) {
			return ExWS.Range[celda1, celda2].Text;
    }

    public List<object> ReaderColumn(int Columna) {
			string R1;
			var swap = new List<object>();
			int X = maxRow(Columna);			
			for (int i = 1; i <= X; i++)
			{
				R1 = Ref2Cel(Columna, i);
				swap.Add(ExWS.Range[R1, R1].Text);
			}
			return swap;
    }

    public List<object> ReaderRow(int Renglon) {
      string R1;
      var swap = new List<object>();
      int X = LastData(Renglon, false, SpaceColumns);
      for (int i = 1; i <= X; i++) {
        R1 = Ref2Cel(i, Renglon);
        swap.Add(ExWS.Range[R1, R1].Text);
      }
      return swap;
    }

    public List<string> ReaderRowStr(int Renglon) {
      string R1;
      var swap = new List<string>();
      int X = LastData(Renglon, false, SpaceColumns);
      for (int i = 1; i <= X; i++) {
        R1 = Ref2Cel(i, Renglon);
        swap.Add(ExWS.Range[R1, R1].Text.ToString());
      }
      return swap;
    }

    public List<object> ReaderRow(int Renglon, int Columna) {
      string R1;
      var swap = new List<object>();
      for (int i = 1; i <= Columna; i++) {
        R1 = Ref2Cel(i, Renglon);
        swap.Add(ExWS.Range[R1, R1].Text);
      }
      return swap;
    }

    #endregion

    #region Exportacion
    public void ExportarDatos(object data, int renglon, int columna) {
      ((Excel.Range)ExWS.Cells[renglon, columna]).Value2 = data;
    }

    public void ExportarDatos(List<object> Data, int renglon, int columna, bool HV) {
      Excel.Range ExRange = null;
      string R1 = string.Empty;
      string R2 = string.Empty;
			if ((Data != null) && (Data.Count > 0))
			{
        if (!HV) {
          R1 = Ref2Cel(columna, renglon);
          R2 = Ref2Cel(Data.Count + columna, renglon);
        }
        else {
          R1 = Ref2Cel(columna, renglon);
          R2 = Ref2Cel(columna, Data.Count + renglon);
        }
        ExRange = ExWS.Range[R1, R2];
        ExRange.Value2 = Data.ToArray();
      }
      else {
        Row++;
        R1 = Ref2Cel(columna, renglon);
        ExRange = ExWS.Range[R1, R1];
        ExRange.Value2 = "No Existen Valores para el rango seleccionado";
      }
      System.Runtime.InteropServices.Marshal.ReleaseComObject(ExRange);
    }

    public void ExportarDatos(List<object> Data, int FormatComas, bool msg) {
      Excel.Range ExRange = null;
      string R1 = string.Empty;
      string R2 = string.Empty;
			if ((Data != null) && (Data.Count > 0))
			{
        if (!Data[0].GetType().IsGenericType) {
          foreach (string s in Data) {
            //Por regla de Consar, se inicia en la columna B
            Row++;
            R1 = Ref2Cel(1, Row);
            R2 = Ref2Cel(1, Row);
            ExRange = ExWS.Range[R1, R2];
          	ExRange.Value2 = s;
          	ExRange.Style = "Comma";
          	ExRange.NumberFormat = FormatoDigitos(FormatComas);
          }
        }
        else {
          foreach (List<object> swap in Data) {
            R1 = Ref2Cel(1, Row);
            R2 = Ref2Cel(swap.Count, Row);
            ExRange = ExWS.Range[R1, R2];
          	ExRange.Value2 = swap.ToArray();
          	ExRange.Style = "Comma";
          	ExRange.NumberFormat = FormatoDigitos(FormatComas);
          	Row++;
          }
        }
      }
      else {
        Row++;
        R1 = Ref2Cel(1, Row);
        R2 = Ref2Cel(1, Row);
        ExRange = ExWS.Range[R1, R2];
        if (msg)
          ExRange.Value2 = "No Existen Valores para el rango seleccionado";
      }
      System.Runtime.InteropServices.Marshal.ReleaseComObject(ExRange);
      //ExRange = null;      
    }

    public void ExportarDatos(List<object> Data, int FormatComas, int Renglon) {
      Excel.Range ExRange = null;
      string R1, R2;
			if ((Data != null) && (Data.Count > 0))
			{
        if (!Data[0].GetType().IsGenericType) {
          int i = Renglon;
          foreach (string s in Data) {
            //Por regla de Consar, se inicia en la columna B          
            R1 = Ref2Cel(1, i);
            R2 = Ref2Cel(1, i);
            ExRange = ExWS.Range[R1, R2];
            ExRange.Value2 = s;
            ExRange.Style = "Comma";
            ExRange.NumberFormat = FormatoDigitos(FormatComas);
            i++;
          }
        }
        else {
          foreach (List<object> swap in Data) {
            R1 = Ref2Cel(1, Row);
            R2 = Ref2Cel(swap.Count, Row);
            ExRange = ExWS.Range[R1, R2];
            ExRange.Value2 = swap.ToArray();
            ExRange.Style = "Comma";
            ExRange.NumberFormat = FormatoDigitos(FormatComas);
            Row++;
          }
        }
      }
      else {
        Row++;
        R1 = Ref2Cel(1, Row);
        R2 = Ref2Cel(1, Row);
        ExRange = ExWS.Range[R1, R2];
        ExRange.Value2 = "No Existen Valores para el rango seleccionado";
      }
      System.Runtime.InteropServices.Marshal.ReleaseComObject(ExRange);
    }

    public void ExportarDatos(List<object> Data, int FormatComas, int Renglon, int Columna) {
      string R1, R2;
			if ((Data != null) && (Data.Count > 0))
			{
        if (!Data[0].GetType().IsGenericType) {
          int i = Renglon;
          foreach (string s in Data) {
            //Por regla de Consar, se inicia en la columna B          
            R1 = Ref2Cel(Columna, i);
            R2 = Ref2Cel(Columna, i);
            ExRange = ExWS.Range[R1, R2];
            ExRange.Value2 = s;
            i++;
          }
        }
        else {
          foreach (List<object> swap in Data) {
            R1 = Ref2Cel(Columna, Row);
            R2 = Ref2Cel(Columna + swap.Count - 1, Row);
            ExRange = ExWS.Range[R1, R2];
            ExRange.Value2 = swap.ToArray();
            Row++;
          }
        }
      }
      else {
        Row++;
        R1 = Ref2Cel(1, Row);
        R2 = Ref2Cel(1, Row);
        ExRange = ExWS.Range[R1, R2];
        ExRange.Value2 = "No Existen Valores para el rango seleccionado";
      }
      System.Runtime.InteropServices.Marshal.ReleaseComObject(ExRange);
    }

    public void ExportarDatos(object[] Data, int FormatComas) {
      int Y;
      string R1, R2;
      Y = Data.GetLength(0);
      R1 = Ref2Cel(1, Row);
      R2 = Ref2Cel(Y, Row);
      ExWS.Range[R1, R2].Value2 = Data;
      ExRange.NumberFormat = FormatoDigitos(FormatComas);
      Row++;
    }
    #endregion

		#region Functions
		public void Graficar(List<object> Data)
		{
			// In projects that target the .NET Framework 3.5, use the following line of code.
			// Worksheet worksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet).GetVstoObject();
			// Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);
			
			var Worksheet = (Tools.Worksheet)ExApp.ActiveSheet;
			Microsoft.Office.Tools.Excel.Chart ExChart;
			ExRange = ExWS.Range["D1", "K16"];
			ExChart = Worksheet.Controls.AddChart(ExRange, "seriesChart");
			ExChart.ChartType = Excel.XlChartType.xlLineStacked;
			ExChart.ChartStyle = 42;
			if ((Data != null) && (Data.Count > 0))
			{
				string R1 = Ref2Cel(1, 1);
				string R2 = Ref2Cel(Data.Count + 1, 1);
				ExRange = ExWS.Range[R1, R2];
				ExRange.Value2 = Data.ToArray();
				ExChart.SetSourceData(ExRange, sinvalor);
			}
		}
		#endregion

		#region Titulos
		public void TitulosCorte(List<object> Titulos) {
      Row++;
      ExEncabezado(ExWS, Row, Titulos);
      Row++;
    }

		private void ExTitulo(Excel.Worksheet ws, string Reporte)
		{
      var configLeyenda = new bool[] { true, true, false, false };
      ExFormato(ws, 1, 1, 1, string.Empty, 16, configLeyenda);
      ExCentrarMix(ExWS, 1, 1, 11);
      var configDatos = new bool[] { true, false, false, false };
      ws.Cells[3, 1] = Reporte;
      ExFormato(ws, 1, 3, 1, string.Empty, 12, configDatos);
      ExCentrarMix(ExWS, 1, 3, 11);
    }

		private void ExTitulo(Excel.Worksheet ws, XmlDocument xReporte)
		{
      XmlNode xNode = xReporte.DocumentElement.SelectSingleNode("Elemento");
      foreach (XmlElement xElement in xNode) {
        int i = xElement.Attributes.Count - 1;
        for (int j = 0; j <= i; j++) {
          ExWS.Range[xElement.Attributes[j].Value, xElement.Attributes[j].Value].Value2 = xElement.Attributes[j + 1].Value;
          j++;
        }
      }
      //ExFormato(ws, xReporte);
    }

		private void ExTitulo(Excel.Worksheet ws, XmlNode xReporte, string Comentario)
		{
      var xElement = (XmlElement)xReporte;
      ExRange = ExWS.Range[xElement.Attributes["Inicial"].Value, xElement.Attributes["Inicial"].Value];
      ExRange.Value2 = xElement.Attributes["Dato"].Value;
      ExRange = ExWS.Range[xElement.Attributes["Inicial"].Value, xElement.Attributes["Final"].Value];
      ExRange.Font.Name = xElement.Attributes["Font"].Value;
      ExRange.Font.Size = xElement.Attributes["Size"].Value;
      ExRange.Font.Bold = xElement.Attributes["Bold"].Value;
      ExRange.Interior.Color = xElement.Attributes["Interior"].Value;
      ExRange.Font.Color = xElement.Attributes["Color"].Value;
      ExRange.Merge(true);
      ExRange.HorizontalAlignment = -4108; //Constante xlCenter
      ExRange.VerticalAlignment = -4107; //Constante xlBottom
      string R1 = Ref2Cel(2, 6);
      ExWS.Range[R1, R1].Value2 = Comentario;
    }

		private void ExEncabezado(Excel.Worksheet ws, int Renglon, object Datos)
		{
      string R1, R2;
      var configHeader = new bool[] { false, false, false, false };
      R1 = Ref2Cel(2, Renglon);
      R2 = Ref2Cel(2, Renglon);
      ExWS.Range[R1, R2].Value2 = Datos;
      ExFormato(ws, 1, Renglon, 1, string.Empty, 8, configHeader);
    }

		private void ExEncabezado(Excel.Worksheet ws, List<object> Datos, string Fondo)
		{
      string R1, R2;
      if (Datos.Count > 0) {
        R1 = Ref2Cel(1, Row);
        R2 = Ref2Cel(Datos.Count, Row);
        ExWS.Range[R1, R2].Value2 = Datos.ToArray();
        ExWS.Range[R1, R2].Interior.Color = Fondo;
        Row++;
      }
    }
    #endregion

    #region Formato
    private string FormatoDigitos(int Coma) {
      string strFormat = string.Empty;
      switch (Coma) {
        case 0:
          strFormat = "_-* #,##0_-;-* #,##0_-;_-* ''-''??_-;_-@_-";
          break;
        case 1:
          strFormat = "_-* #,##0.0_-;-* #,##0.0_-;_-* ''-''??_-;_-@_-";
          break;
        case 2:
          strFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ''-''??_-;_-@_-";
          break;
      }
      return strFormat;
    }

		public void ExFormato(Excel.Worksheet ws, int X, int Y, int Z, string Letra, int Size, bool[] Format)
		{
      string R1, R2;

      R1 = Ref2Cel(X, Y);
      R2 = Ref2Cel(Z, Y);
      ws.Range[R1, R2].Font.Name = Letra == string.Empty ? "Courier New" : Letra;
      ws.Range[R1, R2].Font.Size = Size;
      ws.Range[R1, R2].Font.Bold = Format[0];
      ws.Range[R1, R2].Font.Italic = Format[1];
      ws.Range[R1, R2].Font.Underline = Format[2];
      if (Format[3])
        ws.get_Range(R1, R2).EntireColumn.AutoFit();
    }

		public void ExFormato(Excel.Worksheet ws, XmlDocument xReporte)
		{
      foreach (XmlElement xElement in xReporte.DocumentElement.SelectSingleNode("Formato")) {
        bool inter;
        string R1 = xElement.Attributes["Rango"].Value.Substring(0, 2);
        string R2 = xElement.Attributes["Rango"].Value.Substring(2, 2);
        ws.Range[R1, R2].Font.Name = xElement.Attributes["Letra"].Value;
        ws.Range[R1, R2].Font.Size = xElement.Attributes["Size"].Value;
        bool.TryParse(xElement.Attributes["Bold"].Value, out inter);
        ws.Range[R1, R2].Font.Bold = inter;
        bool.TryParse(xElement.Attributes["Italic"].Value, out inter);
        ws.Range[R1, R2].Font.Italic = inter;
        bool.TryParse(xElement.Attributes["Underline"].Value, out inter);
        ws.Range[R1, R2].Font.Underline = inter;
      }
    }

		public void ExFormato(Excel.Worksheet ws, int X, int Y, int Z, int W, byte Fondo, byte Letra)
		{
      string R1, R2;

      R1 = Ref2Cel(X, Y);
      R2 = Ref2Cel(Z, Y);
      ws.Range[R1, R2].Interior.ColorIndex = Fondo;
      ws.Range[R1, R2].Font.ColorIndex = Letra;
    }

		private void ExCentrarMix(Excel.Worksheet ws, byte X, byte Y, byte Z)
		{
      string R1, R2;

      R1 = Ref2Cel(X, Y);
      R2 = Ref2Cel(Z, Y);
      ws.get_Range(R1, R2).HorizontalAlignment = -4108; //Constante xlCenter
      ws.get_Range(R1, R2).VerticalAlignment = -4107; //Constante xlBottom
      ws.get_Range(R1, R2).WrapText = false;
      ws.get_Range(R1, R2).Orientation = 0;
      ws.get_Range(R1, R2).AddIndent = false;
      ws.get_Range(R1, R2).ShrinkToFit = true;
      ws.get_Range(R1, R2).MergeCells = true;
    }
    #endregion

    #region Rango
    private string Ref2Cel(int Col, int Ren) {
      string S1 = string.Empty;

      if (EnRango(Col, 26, 1))
        S1 = System.Convert.ToChar(System.Convert.ToInt32('A') + Col - 1) + Ren.ToString();
      else
        if (EnRango(Col, 26, 2))
          S1 = "A" + System.Convert.ToChar(System.Convert.ToInt32('A') + Col - 27) + Ren.ToString();
        else
          if (EnRango(Col, 26, 3))
            S1 = "C" + System.Convert.ToChar(System.Convert.ToInt32('B') + Col - 53) + Ren.ToString();
          else
            if (EnRango(Col, 26, 4))
              S1 = "C" + System.Convert.ToChar(System.Convert.ToInt32('C') + Col - 79) + Ren.ToString();
      if (S1 == string.Empty)
        S1 = "A1";
      return S1;
    }

    private int LastData(int Posicion, bool RowCol, int salto) {
      int X, Y;
      string R1, R2;
      if (RowCol) {  //Se obtiene el ultimo registro horizontal
        Y = salto;
        X = Posicion == 0 ? 1 : Posicion;
        R1 = Ref2Cel(X, Y);
        while (ExWS.Range[R1, R1].Text.ToString() != string.Empty) {
          Y = Y + salto - 1;
          R1 = Ref2Cel(X, Y);
        }
        while (ExWS.Range[R1, R1].Text.ToString() == string.Empty) {
          Y--;
          R1 = Ref2Cel(X, Y);
        }
        return Y;
      }
    	//Se obtiene el ultimo registro vertical
    	X = salto;
    	Y = Posicion == 0 ? 1 : Posicion;
    	R2 = Ref2Cel(X, Y);
    	string obj = ExWS.Range[R2, R2].Text.ToString();
    	while (ExWS.Range[R2, R2].Text.ToString() != string.Empty) {
    		X = X + salto - 1;
    		R2 = Ref2Cel(X, Y);
    	}
    	while (ExWS.Range[R2, R2].Text.ToString() == string.Empty) {
    		X--;
    		R2 = Ref2Cel(X, Y);
    	}
    	return X;
    }
    
    private bool EnRango(int Numero, int Salto, int Fase)
    {
    	return (Numero >= (Salto * (Fase - 1) + 1)) && (Numero <= (Salto * (Fase)));
    }
  	#endregion

    #region Finaliza
    public void Salvar(string Archivo) {
      try {
        if (File.Exists(Archivo))
          ExWB.Save();
        ExWB.SaveAs(Archivo, Excel.XlFileFormat.xlXMLSpreadsheet, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
      	Cerrar(Archivo);
      }
      catch (Exception ex) {
        throw new Exception(ex.Message + "Error en Excel - No se puede guardar el Archivo", ex);
      }
    }

    public void Cerrar(string Archivo) {
      try {
        ExWB.Close(true, Archivo, false);
				ExApp.ActiveWorkbook.Close(false, false, sinvalor);
				Finalizar();
      }
      catch (Exception ex) {
        throw new Exception(ex.Message + "Error en Excel - No se puede cerrar el Archivo", ex);
      }
    }

    public void Termina() {
      try {
				ExApp.ActiveWorkbook.Close(false, false, sinvalor);
      }
      catch (Exception ex) {
        throw new Exception(ex.Message + "Error en Excel - No se puede guardar el Archivo", ex);
      }
      finally {
        Finalizar();
      }
    }

    public void Finalizar() {
      try {
        ExApp.Quit();

				if (ExRange != null) Marshal.ReleaseComObject(ExRange);
				if (ExWS != null) Marshal.ReleaseComObject(ExWS);
				if (ExS != null) Marshal.ReleaseComObject(ExS);
				if (ExWB != null) Marshal.ReleaseComObject(ExWB);
				if (ExApp != null) Marshal.ReleaseComObject(ExApp);

				ExRange = null;
				ExWS = null;
				ExS = null;
				ExWB = null;
				ExApp = null;

				GC.Collect();
				GC.WaitForPendingFinalizers();
				GC.Collect();
      }
			catch (Exception ex)
			{
				throw new Exception(ex.Message, ex);
			}
    }

		// Cierra archivo de excel en uso		
		public void CerrarArchivoExcel()
		{
			System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
			ExApp.ActiveWorkbook.Close(false, false, sinvalor);
		}

		// Cierra la aplicación de excel en uso            
		public void CerrarExcel()
		{
			System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
			try
			{
				ExApp.Visible = false;
				ExApp.UserControl = false;
				ExWB.Close(false, false, null);
			}
			catch (Exception ex)
			{
				throw new Exception(ex.Message, ex);
			}
			Finalizar();
		}
    #endregion

    #region XML
    public XmlDocument Xls2Xml(List<string> Campos, int InitRow) {
      StringBuilder sb = new StringBuilder();
      XmlDocument xDoc = new XmlDocument();
      sb.Append("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");  // se crea la declaracion de XML
      sb.Append(" <QryName> ");	// se adiciona la declaración de XML al documento XML
      try {
        for (int i = InitRow; i <= maxRow(1); i++) {
          sb.Append("<detail ");
          for (int j = 1; j <= Campos.Count; j++) {
            string Celda = ReaderCell(j, i).ToString();
            string CeldaProcesada = Celda.Replace("<", "&lt;").Replace(">", "&gt;");
            sb.Append(Campos[j - 1] + "='" + CeldaProcesada + "' ");
          }
          sb.Append(" />");
          sb.Append("</Excel>");
        }
        sb.Append(" </QryName> ");
        string xml = sb.ToString();
        xDoc.LoadXml(xml);
      }
      catch (Exception ex) {
        throw new Exception(ex.Message, ex);
      }
      return xDoc;
    }

    public XmlDocument Xls2Xml(int InitRow) {
      var sb = new StringBuilder();
      var xDoc = new XmlDocument();
      sb.Append("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");  // se crea la declaracion de XML
      sb.Append(" <QryName> ");	// se adiciona la declaración de XML al documento XML
      try {
        for (int i = InitRow; i <= maxRow(1); i++) {
          sb.Append("<Excel>");
          for (int j = 1; j <= ReaderRowStr(5).Count; j++) {
            object cell = ((Excel.Range)ExWS.Cells[i, j]).Value2;
            if (cell == null) {
              sb.Append("<detail valor='' /> ");
            }
            else {
              string CeldaProcesada = cell.ToString().Replace("<", "&lt;").Replace(">", "&gt;");
              sb.Append("<detail valor='" + CeldaProcesada + "' /> ");
            }
          }
          sb.Append("</Excel>");
        }
        sb.Append(" </QryName> ");
        string xml = sb.ToString();
        xDoc.LoadXml(xml);
      }
      catch (Exception ex) {
        throw new Exception(ex.Message, ex);
      }
      return xDoc;
    }

		public XmlDocument Xls2Xml(int rowInicial, int rowFinal, int columna)
		{
			var sb = new StringBuilder();
			var xDoc = new XmlDocument();
			sb.Append("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");  // se crea la declaracion de XML
			sb.Append(" <Conversion> ");	// se adiciona la declaración de XML al documento XML
			try
			{
				for (int i = rowInicial; i <= rowFinal; i++)
				{
					sb.Append("<Excel>");
					for (int j = 1; j <= columna; j++)
					{
						object cell = (ExWS.Cells[i, j]).Value2;
						if (cell == null)
							sb.AppendFormat("<columna{0}/>", j);
						else
						{
							string celdaProcesada = cell.ToString().Replace("<", "&lt;").Replace(">", "&gt;");
							sb.AppendFormat("<columna{0}>{1}</columna{0}>", j, celdaProcesada);
						}
					}
					sb.Append("</Excel>");
				}
				sb.Append(" </Conversion> ");
				string xml = sb.ToString();
				xDoc.LoadXml(xml);
			}
			catch (Exception ex)
			{
				throw new Exception(ex.Message, ex);
			}
			return xDoc;
		}
    #endregion

    public void Xls2Csv(string sourceFile, string targetFile) {
      ExWB = ExApp.Workbooks.Open(sourceFile, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
      ExWB.SaveAs(targetFile, Excel.XlFileFormat.xlCSVWindows, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
      ExWB.Close(true, targetFile, false);
    }
  }
}