using System.Runtime.InteropServices;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.Versioning;
using System.Security;



namespace ExcelMethods
{
    public class ExcelClass
    {
        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------

        // Escribir todos los atributos de clase necesarios para los métodos aquí.
        // Añadir también una descripción de cada uno para poder localizarlos.

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
        /// A partir de una carpeta con archivos de SAP, se obtienen las rutas de guardado 
        /// de los excels en la carpeta que queramos usando el mismo nombre que los ficheros 
        /// .sdb . Por ejemplo, si tenemos una carpeta con ficheros SAP y queremos guardar las 
        /// reacciones de los mismos y que cada excel tenga el mismo nombre que el .sdb, se 
        /// dará esa carpeta de archivos SAP, una lista de strings de las rutas de los ficheros 
        /// .sdb, y la carpeta donde se desea guardar los excels.
        /// </summary>
        /// <param name="SAPFilesroutes">
        /// Lista de strings con las rutas de los ficheros de SAP. 
        /// </param>
        /// <param name="SAPFolderRoute">
        /// Ruta de la carpeta donde se encuentran los archivos de SAP (string). 
        /// </param>
        /// <param name="ExcelFolderRoute">
        /// Ruta de la carpeta donde se desean guardar los excel (string). 
        /// </param>
        /// <returns>Lista de strings con las rutas de guardado de todos los excel.</returns>
        public List<string> EstablishExcelRoutes(List<string> SAPFilesRoutes, string SAPFolderRoute, string ExcelFolderRoute)
        {
            List<string> ExcelFilesRoutes = new List<string>();

            foreach (string route in SAPFilesRoutes)
            {
                string ExcelRoute = route.Replace(SAPFolderRoute, ExcelFolderRoute);
                ExcelRoute = System.IO.Path.ChangeExtension(ExcelRoute, ".xlsx");
                ExcelFilesRoutes.Add(ExcelRoute);
            }

            return ExcelFilesRoutes;
        }


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


        //---------------------------------------------------------------------------------



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