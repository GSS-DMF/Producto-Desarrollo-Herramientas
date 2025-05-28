using System.Runtime.InteropServices;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.Versioning;
using System.Security;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using System.Diagnostics;




namespace RepositorioFuncionesGitHub
{
    public class MSExcel
    {
        static MSExcel() // Ponemos la licencia no comercial para permitir usar EPPLUS
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }


        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------

        // Escribir todas las instancias de subclases aquí. 

        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------



        public MSExcel()
        {
            // Constructor de la clase Excel
            Format = new FormatSubclass(this);
        }

        public FormatSubclass Format { get; }



        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------

        // Traemos las propiedades de clase de la clase pricipal

        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------




        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------

        // Escribir todos los atributos de clase necesarios para los métodos aquí.
        // Añadir también una descripción de cada uno para poder localizarlos. Añadirlos 
        // todos como propiedades static para que las subclases tengan acceso a ellas.

        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------







        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------

        // Escribir todos los métodos de clase aquí. Añadir un docstring para
        // tener información acerca de su funcionamiento y parámetros de entrada
        // y salida. Recordar añadirlo al excel de registro de métodos. Poner 
        // todos los métodos públicos para evitar errores de acceso.

        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------


        /// <summary>
        /// Toma un excel abierto por SAP2000 (que no está guardado en ninguna ruta, por eso 
        /// inexistente) y lo trata de guardar en la ruta que le pasamos. NOTA: a veces no 
        /// funciona correctamente, y la etiqueta de confidencialidad se debe poner de forma 
        /// manual.
        /// </summary>
        /// <param name="ExcelFileRoute">
        /// Ruta donde se desea guardar el archivo excel (string). 
        /// </param>
        public void SaveInexistentExcel(string ExcelFileRoute)
        {
            try
            {
                // Intentar obtener la aplicación de Excel abierta
                Excel.Application excelApp = null;

                try
                {
                    excelApp = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
                }
                catch (COMException)
                {
                    MessageBox.Show("No hay una instancia activa de Excel.");
                    return;
                }

                // Verificar si hay libros abiertos
                if (excelApp.Workbooks.Count == 0)
                {
                    MessageBox.Show("No hay libros abiertos en Excel.");
                    return;
                }

                // Tomar el libro activo
                Excel.Workbook libro = excelApp.ActiveWorkbook;

                if (libro != null)
                {
                    // Guardar el libro en la ruta especificada
                    libro.SaveAs(ExcelFileRoute, Excel.XlFileFormat.xlOpenXMLWorkbook);
                    libro.Close(false);
                    libro = null;
                }
                else
                {
                    MessageBox.Show("No se pudo encontrar el archivo generado por SAP2000.");
                }

                // Cerrar la aplicación de Excel si se necesita
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                excelApp = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al guardar el archivo de Excel: " + ex.Message);
            }
        }

        /// <summary>
        /// Toma un excel abierto por SAP2000 (que no está guardado en ninguna ruta, por eso 
        /// inexistente) y lo cierra.
        /// </summary>
        public void CloseExcel()
        {
            try
            {
                // Intentar obtener la aplicación de Excel abierta
                Excel.Application excelApp = null;

                try
                {
                    excelApp = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
                }
                catch (COMException)
                {
                    MessageBox.Show("No hay una instancia activa de Excel.");
                    return;
                }

                // Verificar si hay libros abiertos
                if (excelApp.Workbooks.Count == 0)
                {
                    MessageBox.Show("No hay libros abiertos en Excel.");
                    return;
                }

                // Tomar el libro activo
                Excel.Workbook libro = excelApp.ActiveWorkbook;

                if (libro != null)
                {
                    // Cerrar el libro
                    libro.Close(false);
                    Marshal.ReleaseComObject(libro);
                    libro = null;
                }
                else
                {
                    MessageBox.Show("No se pudo encontrar el archivo generado por SAP2000.");
                }

                // Cerrar la aplicación de Excel si se necesita
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                excelApp = null;

                // Forzar la recolección de basura para liberar los objetos COM
                GC.Collect();
                GC.WaitForPendingFinalizers();
                TerminateExcelProcesses();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al guardar el archivo de Excel: " + ex.Message);
            }
        }

        ///<summary>
        ///Intenta obtener un libro abierto de excel que no esté guardado en ninguna ruta
        ///</summary>
        ///<returns>
        ///Devuelve la ruta del excel si lo encuentra, si no devuelve null
        ///</returns>
        public Excel.Workbook CatchExcel()
        {
            Excel.Application excelApp = null;
            Excel.Workbook libro = null;

            try
            {
                // Intentar obtener la aplicación de Excel abierta
                try
                {
                    excelApp = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
                }
                catch (COMException)
                {
                    MessageBox.Show("No hay una instancia activa de Excel.");
                    return null;
                }

                // Verificar si hay libros abiertos
                if (excelApp.Workbooks.Count == 0)
                {
                    MessageBox.Show("No hay libros abiertos en Excel.");
                    return null;
                }

                // Tomar el libro activo
                libro = excelApp.ActiveWorkbook;

                return libro;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al guardar el archivo de Excel: " + ex.Message);
                return null;
            }

        }

        ///<summary>
        ///Obtener todos los procesos de Excel en ejecución y finalizarlos
        /// </summary>
        public void TerminateExcelProcesses()
        {
            // Obtener todos los procesos de Excel en ejecución
            Process[] excelProcesses = Process.GetProcessesByName("Excel");

            foreach (Process process in excelProcesses)
            {
                try
                {
                    // Finalizar el proceso de Excel
                    process.Kill();
                    process.WaitForExit(); // Esperar a que el proceso termine
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error al finalizar el proceso de Excel: " + ex.Message);
                }
            }
        }

        ///<summary>
        ///Elimina la fila indicada de un objeto excel abierto
        ///</summary>
        ///<param name="row">
        ///Valor de la fila que se quiere eliminar del excel
        ///</param>
        ///<param name="libro">
        ///Objeto libro abierto
        ///</param>
        public void DeleteExcelTableRow(int row, Excel.Workbook libro)
        {
            try
            {
                if (libro != null)
                {
                    //Tomar la hoja activa
                    Excel.Worksheet hoja=(Microsoft.Office.Interop.Excel.Worksheet)libro.ActiveSheet;

                    //Eliminar fila indicada
                    Excel.Range fila = (Microsoft.Office.Interop.Excel.Range)hoja.Rows[row];
                    fila.Delete();

                    // Liberar objetos COM
                    Marshal.ReleaseComObject(hoja);
                    hoja = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al guardar el archivo de Excel: " + ex.Message);
            }
        }

        ///<summary>
        ///Mantiene las columnas deseadas de una tabla de un objeto excel abierto
        ///</summary>
        ///<param name="columnNames">
        ///Nombres de la columnas que se desea mantener
        ///</param>
        ///<param name="libro">
        ///Objeto excel abierto
        ///</param>
        public void KeepExcelTableColumns(string[] columnNames, Excel.Workbook libro)
        {
            try
            {
                if (libro != null)
                {
                    //Tomar la hoja activa
                    Excel.Worksheet hoja = (Microsoft.Office.Interop.Excel.Worksheet)libro.ActiveSheet;

                    //Eliminar la columna no deseada
                    Excel.Range usedRange = hoja.UsedRange;
                    for (int i = usedRange.Columns.Count; i >= 1; i--)
                    {
                        Excel.Range column = (Excel.Range)usedRange.Columns[i];
                        Excel.Range cell = (Excel.Range)column.Cells[1, 1]; // 1 fila para nombres de columnas
                        string nombreColumnaExcel = cell.Value2 != null ? cell.Value2.ToString() : string.Empty;

                        // Eliminar la columna si el nombre coincide con el columnName
                        for (int j = 0; j < columnNames.Length; j++)
                        {
                            if (nombreColumnaExcel.Equals(columnNames[i], StringComparison.OrdinalIgnoreCase))
                            {
                                column.Delete();
                                break; // Salir del bucle después de eliminar la columna
                            }
                        }
                    }

                    // Liberar objetos COM
                    Marshal.ReleaseComObject(hoja);
                    hoja = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al guardar el archivo de Excel: " + ex.Message);
            }
        }

        ///<summary>
        ///Filtrar tabla de un objeto excel abierto, según si los valores especificados
        ///coinciden con los de una columna dada
        ///</summary>
        ///<param name="columnaFiltro">
        ///Columna a filtrar
        ///</param>
        ///<param name="valorFiltro">
        ///Valor por el que se va a filtrar la columna
        ///</param>
        ///<param name="libro">
        ///Objeto excel abierto
        ///</param>
        public void FilterTableEqual(string columnaFiltro, string valorFiltro, Excel.Workbook libro)
        {
            try
            {
                if (libro != null)
                {
                    Excel.Worksheet hoja = (Microsoft.Office.Interop.Excel.Worksheet)libro.ActiveSheet;

                    // Encontrar el índice de la columna basada en el nombre
                    Excel.Range filaEncabezado = (Microsoft.Office.Interop.Excel.Range)hoja.Rows[1];
                    int numerocolumnaFiltro = -1;
                    for (int i = 1; i <= filaEncabezado.Columns.Count; i++)
                    {
                        Excel.Range cell = (Excel.Range)filaEncabezado.Cells[1, i];
                        if (cell.Value2 != null && cell.Value2.ToString() == columnaFiltro)
                        {
                            numerocolumnaFiltro = i;
                            break;
                        }
                        Marshal.ReleaseComObject(cell);
                    }

                    if (numerocolumnaFiltro == -1)
                    {
                        MessageBox.Show("No se encontró la columna especificada.");
                        return;
                    }

                    hoja.UsedRange.AutoFilter(numerocolumnaFiltro, valorFiltro, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al guardar el archivo de Excel: " + ex.Message);
            }
        }

        ///<summary>
        ///Filtrar tabla de un objeto excel abierto según si los valores especificados 
        ///no coinciden con los de una columna dada
        ///</summary>
        ///<param name="columnaFiltro">
        ///Columna a filtrar
        ///</param>
        ///<param name="valorFiltro">
        ///Valor por el que se va a filtrar la columna
        ///</param>
        ///<param name="libro">
        ///Objeto excel abierto
        ///</param>
        public void FilterTableNotEqual(string columnaFiltro, string valorFiltro, Excel.Workbook libro)
        {
            try
            {
                if (libro != null)
                {
                    Excel.Worksheet hoja = (Microsoft.Office.Interop.Excel.Worksheet)libro.ActiveSheet;

                    // Encontrar el índice de la columna basada en el nombre
                    Excel.Range filaEncabezado = (Microsoft.Office.Interop.Excel.Range)hoja.Rows[1];
                    int numerocolumnaFiltro = -1;
                    for (int i = 1; i <= filaEncabezado.Columns.Count; i++)
                    {
                        Excel.Range cell = (Excel.Range)filaEncabezado.Cells[1, i];
                        if (cell.Value2 != null && cell.Value2.ToString() == columnaFiltro)
                        {
                            numerocolumnaFiltro = i;
                            break;
                        }
                        Marshal.ReleaseComObject(cell);
                    }

                    if (numerocolumnaFiltro == -1)
                    {
                        MessageBox.Show("No se encontró la columna especificada.");
                        return;
                    }

                    hoja.UsedRange.AutoFilter(numerocolumnaFiltro, "<>" + valorFiltro, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al guardar el archivo de Excel: " + ex.Message);
            }
        }

        ///<summary>
        ///Filtrar tabla de un objeto excel abierto según si los valores especificados son mayores o menores que
        ///los de una columna dada
        ///</summary>
        ///<param name="columnaFiltro">
        ///Columna a filtrar
        ///</param>
        ///<param name="valorFiltro">
        ///Valor por el que se va a filtrar la columna
        ///</param>
        ///<param name="libro">
        ///Objeto excel abierto
        ///</param>
        ///<param name="minor">
        ///Variable opcional para elegir entre "mayor que" o "menor que". Por defecto el valor es false, 
        ///por lo que la función compararía con "mayor que"
        ///</param>
        public void FilterTableByComparison(string columnaFiltro, string valorFiltro, Excel.Workbook libro, bool? minor = null)
        {
            try
            {
                if (libro != null)
                {
                    Excel.Worksheet hoja = (Microsoft.Office.Interop.Excel.Worksheet)libro.ActiveSheet;

                    // Encontrar el índice de la columna basada en el nombre
                    Excel.Range filaEncabezado = (Microsoft.Office.Interop.Excel.Range)hoja.Rows[1];
                    int numerocolumnaFiltro = -1;
                    for (int i = 1; i <= filaEncabezado.Columns.Count; i++)
                    {
                        Excel.Range cell = (Excel.Range)filaEncabezado.Cells[1, i];
                        if (cell.Value2 != null && cell.Value2.ToString() == columnaFiltro)
                        {
                            numerocolumnaFiltro = i;
                            break;
                        }
                        Marshal.ReleaseComObject(cell);
                    }

                    if (numerocolumnaFiltro == -1)
                    {
                        MessageBox.Show("No se encontró la columna especificada.");
                        return;
                    }

                    if (minor == true)
                    {
                        hoja.UsedRange.AutoFilter(numerocolumnaFiltro, "<" + valorFiltro, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
                    }
                    else if (minor == false)
                    {
                        hoja.UsedRange.AutoFilter(numerocolumnaFiltro, ">" + valorFiltro, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al guardar el archivo de Excel: " + ex.Message);
            }
        }

        ///<summary>
        ///Copiar el rango activo de un archivo de un excel abierto
        ///</summary>
        ///<param name="libro">
        ///Objeto excel abierto
        ///</param>
        public void CopyExcel(Excel.Workbook libro)
        {
            try
            {
                if (libro != null)
                {
                    Excel.Worksheet hoja = (Microsoft.Office.Interop.Excel.Worksheet)libro.ActiveSheet;
                    // Contar las celdas no vacías en la primera fila después del filtrado
                    int nonEmptyCellCount = 0;
                    Excel.Range firstRow = (Microsoft.Office.Interop.Excel.Range)hoja.Rows[1];
                    for (int i = 1; i <= 40; i++)
                    {
                        Excel.Range cell = (Microsoft.Office.Interop.Excel.Range)firstRow.Cells[1, i];
                        if (cell.Value2 != null && !string.IsNullOrEmpty(cell.Value2.ToString()))
                        {
                            nonEmptyCellCount++;
                        }
                    }

                    // Copiar solo las columnas no vacías
                    Excel.Range range = hoja.Range[hoja.Cells[1, 1], hoja.Cells[hoja.UsedRange.Rows.Count, nonEmptyCellCount]];
                    range.Copy();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al guardar el archivo de Excel: " + ex.Message);
            }
        }

        ///<summary>
        ///Abre excel en segundo plano
        ///</summary>
        public void StartExcel()
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet sheet=(Excel.Worksheet)workbook.Sheets[1];
        }

        /// <summary>
        /// Pega una tabla en un libro excel nuevo que abre en segundo plano. Se puede utilizar para
        /// trabajar con tablas en segundo plano de una forma rápida
        /// </summary>
        /// <param name="table">
        /// Tabla de datos que se quiere copiar
        /// </param>
        /// <returns>
        /// Devuelve el objeto libro de excel que ha abierto para copiar la tabla
        /// </returns>
        public Excel.Workbook PasteTableInExcel(string[,] table)
        {
            Excel.Workbook workbook = null;

            try
            {
                int numFilas = table.GetLength(0);
                int numColumnas = table.GetLength(1);

                //Iniciar Excel
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = false;
                workbook = excelApp.Workbooks.Add();
                Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[1];

                //Crear un array 2D para asignar de una sola vez
                object[,] excelData=new object[numFilas,numColumnas];
                for (int i = 0; i < numFilas; i++)
                {
                    for (int j = 0; j < numColumnas; j++)
                    {
                        excelData[i,j] = table[i,j];
                    }
                }

                // Asignar el rango de una sola vez
                Excel.Range startCell = (Excel.Range)sheet.Cells[1, 1];
                Excel.Range endCell = (Excel.Range)sheet.Cells[numFilas, numColumnas];
                Excel.Range writeRange = sheet.Range[startCell, endCell];
                writeRange.Value2 = excelData;

                return workbook;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al copiar la tabla en Excel: " + ex.Message);
                return null;
            }
        }



        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------

        // Escribir todas las subclases aquí. Añadir un docstring para tener 
        // información acerca de su funcionamiento y parámetros de entrada
        // y salida. Recordar añadirlo al excel de registro de métodos. Poner 
        // todos los métodos públicos para evitar errores de acceso.

        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------



        public class FormatSubclass
        {
            private readonly MSExcel _msexcel;

            public FormatSubclass(MSExcel msexcel)
            {
                _msexcel = msexcel;
            }



            ///<summary>
            ///Dar formato a la fuente de un objeto de excel abierto
            ///</summary>
            ///<param name="fontName">
            ///Nombre de la fuente a asignar
            ///</param>
            ///<param name="fontSize">
            ///Valor del tamaño de fuent a asignar
            ///</param>
            ///<param name="libro">
            ///Objeto excel abierto
            ///</param>
            ///<param name="horizontalAlignmentCenter">
            ///Variable que activa centrar horizontalmente el contenido de las celdas de todo el archivo
            ///Por defecto aparece desactivado.
            ///</param>
            ///<param name="verticalAlignmentCenter">
            ///Variable que activa centrar verticalmente el contenido de las celdas de todo el archivo
            ///Por defecto aparece desactivado.
            ///</param>
            public void ApplyFont(int fontSize, string fontName, Excel.Workbook libro, bool? horizontalAlignmentCenter = false, bool? verticalAlignmentCenter = false)
            {
                try
                {
                    if (libro != null)
                    {
                        //Tomar la hoja activa
                        Excel.Worksheet hoja = (Microsoft.Office.Interop.Excel.Worksheet)libro.ActiveSheet;

                        //Dar formato
                        hoja.Cells.Font.Name = fontName;
                        hoja.Cells.Font.Size = fontSize;
                        if (horizontalAlignmentCenter == true)
                        {
                            hoja.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                        if (verticalAlignmentCenter == true)
                        {
                            hoja.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        }

                        // Liberar objetos COM
                        Marshal.ReleaseComObject(hoja);
                        hoja = null;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al guardar el archivo de Excel: " + ex.Message);
                }
            }

            ///<summary>
            ///Dar formato de color a la fuente de un objeto excel abierto
            ///</summary>
            ///<param name="fontColor">
            ///Color a asignar a la fuente de todo el archivo
            ///</param>
            ///<param name="interiorColor">
            ///Color de fondo a asignar a la fuente de todo el archivo
            ///</param>
            ///<param name="libro">
            ///Objeto excel abierto
            ///</param>
            public void ApplyColor(Color fontColor, Color interiorColor, Excel.Workbook libro)
            {
                try
                {
                    if (libro != null)
                    {
                        //Tomar la hoja activa
                        Excel.Worksheet hoja = (Microsoft.Office.Interop.Excel.Worksheet)libro.ActiveSheet;

                        //Dar formato
                        hoja.Cells.Font.Color = ColorTranslator.ToOle(fontColor);
                        hoja.Cells.Interior.Color = ColorTranslator.ToOle(interiorColor);

                        // Liberar objetos COM
                        Marshal.ReleaseComObject(hoja);
                        hoja = null;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al guardar el archivo de Excel: " + ex.Message);
                }
            }

            ///<summary>
            ///Dar formato a la fuente de una fila de hoja de excel abierta
            ///</summary>
            ///<param name="fontName">
            ///Nombre de la fuente a asignar
            ///</param>
            ///<param name="fontSize">
            ///Valor del tamaño de fuent a asignar
            ///</param>
            ///<param name="row">
            ///Fila a la que se le va a asignar el formato
            ///</param>
            ///<param name="libro">
            ///Objeto excel abierto
            ///</param>
            ///<param name="horizontalAlignmentCenter">
            ///Variable que activa centrar horizontalmente el contenido de las celdas de todo el archivo
            ///Por defecto aparece desactivado.
            ///</param>
            ///<param name="verticalAlignmentCenter">
            ///Variable que activa centrar verticalmente el contenido de las celdas de todo el archivo
            ///Por defecto aparece desactivado.
            ///</param>
            public void ApplyFontToRow(int fontSize, string fontName, int row, Excel.Workbook libro, bool? horizontalAlignmentCenter = false, bool? verticalAlignmentCenter = false)
            {
                try
                {
                    if (libro != null)
                    {
                        //Tomar la hoja activa
                        Excel.Worksheet hoja = (Microsoft.Office.Interop.Excel.Worksheet)libro.ActiveSheet;

                        //Dar formato a la fila
                        Excel.Range fila = (Microsoft.Office.Interop.Excel.Range)hoja.Rows[row];
                        fila.Cells.Font.Name = fontName;
                        fila.Cells.Font.Size = fontSize;
                        if (horizontalAlignmentCenter == true)
                        {
                            fila.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                        if (verticalAlignmentCenter == true)
                        {
                            fila.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        }

                        // Liberar objetos COM
                        Marshal.ReleaseComObject(hoja);
                        hoja = null;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al guardar el archivo de Excel: " + ex.Message);
                }
            }

            ///<summary>
            ///Dar formato a la fuente de una fila de una hoja de excel abierta
            ///</summary>
            ///<param name="row">
            ///Fila a la que se le va a aplicar el formato
            ///</param>
            ///<param name="fontColor">
            ///Color a asignar a la fuente de todo el archivo
            ///</param>
            ///<param name="interiorColor">
            ///Color de fondo a asignar a la fuente de todo el archivo
            ///</param>
            ///<param name="libro">
            ///Objeto excel abierto
            ///</param>
            public void ApplyColorToRow(int row, Color fontColor, Color interiorColor, Excel.Workbook libro)
            {
                try
                {
                    if (libro != null)
                    {
                        //Tomar la hoja activa
                        Excel.Worksheet hoja = (Microsoft.Office.Interop.Excel.Worksheet)libro.ActiveSheet;

                        //Dar formato a la fila
                        Excel.Range fila = (Microsoft.Office.Interop.Excel.Range)hoja.Rows[row];
                        fila.Cells.Font.Color = ColorTranslator.ToOle(fontColor);
                        fila.Cells.Interior.Color = ColorTranslator.ToOle(interiorColor);

                        // Liberar objetos COM
                        Marshal.ReleaseComObject(hoja);
                        hoja = null;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al guardar el archivo de Excel: " + ex.Message);
                }
            }

            ///<summary>
            ///Dar formato de encabezado de las tablas de reacciones que saca SAP2000 
            ///a un excel que queramos en un rango dado. Solo pone el formato, no el texto.
            ///</summary>
            ///<param name="ExcelFileRoute">
            ///Excel al que aplicar el formato
            ///</param>
            ///<param name="Sheet">
            ///Hoja del excel en la que aplicar el formato
            ///</param>
            ///<param name="PasteRange">
            ///Primera celda en la que aplicar el formato (celda de arriba la izquierda)
            ///</param>
            public void AddSAPReactionsHeader(string ExcelFileRoute, string Sheet, string PasteRange)
            {
                FileInfo archivoExcel = new FileInfo(ExcelFileRoute);

                using (ExcelPackage paquete = new ExcelPackage(archivoExcel))
                {
                    ExcelWorksheet hoja = paquete.Workbook.Worksheets[Sheet];
                    var inicio = hoja.Cells[PasteRange].Start;

                    int startRow = inicio.Row;
                    int startCol = inicio.Column;
                    int columnas = 10;

                    var rangoFila1 = hoja.Cells[startRow, startCol, startRow, startCol + columnas - 1];
                    rangoFila1.Merge = true;
                    rangoFila1.Style.Font.Bold = true;
                    rangoFila1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rangoFila1.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(51, 204, 204));

                    for (int i = 0; i < columnas; i++)
                    {
                        var celda = hoja.Cells[startRow + 1, startCol + i];
                        celda.Style.Font.Bold = true;
                        celda.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        celda.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 255, 255));
                    }

                    for (int i = 0; i < columnas; i++)
                    {
                        var celda = hoja.Cells[startRow + 2, startCol + i];
                        celda.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        celda.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 255, 255));
                    }

                    paquete.Save();
                }
            }
        }
    }




    // Clase auxiliar para poder llamar al método Marshal.GetActiveObject, sino no funciona guardar
    // excels. Es un tipo de actualización al Marshal que liberó Microsoft pero no a todas las 
    // versiones. A la nuestra no, así que por eso lo incluimos. NO MODIFICAR.

    public static class Marshal2
    {
        internal const String OLEAUT32 = "oleaut32.dll";
        internal const String OLE32 = "ole32.dll";

        [System.Security.SecurityCritical]  // auto-generated_required
        public static Object GetActiveObject(String progID)
        {
            Object obj = null;
            Guid clsid;

            // Call CLSIDFromProgIDEx first then fall back on CLSIDFromProgID if
            // CLSIDFromProgIDEx doesn't exist.
            try
            {
                CLSIDFromProgIDEx(progID, out clsid);
            }
            //            catch
            catch (Exception)
            {
                CLSIDFromProgID(progID, out clsid);
            }

            GetActiveObject(ref clsid, IntPtr.Zero, out obj);
            return obj;
        }

        //[DllImport(Microsoft.Win32.Win32Native.OLE32, PreserveSig = false)]
        [DllImport(OLE32, PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [System.Security.SecurityCritical]  // auto-generated
        private static extern void CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] String progId, out Guid clsid);

        //[DllImport(Microsoft.Win32.Win32Native.OLE32, PreserveSig = false)]
        [DllImport(OLE32, PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [System.Security.SecurityCritical]  // auto-generated
        private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] String progId, out Guid clsid);

        //[DllImport(Microsoft.Win32.Win32Native.OLEAUT32, PreserveSig = false)]
        [DllImport(OLEAUT32, PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [System.Security.SecurityCritical]  // auto-generated
        private static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.Interface)] out Object ppunk);
    }
}