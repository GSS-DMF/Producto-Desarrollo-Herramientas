using SAP2000v1;
using System.IO;


namespace SAPMethods
{
    public class SAPClass
    {
        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------

        // Escribir todos los atributos de clase necesarios para los métodos aquí.
        // Añadir también una descripción de cada uno para poder localizarlos.

        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------


        cOAPI mySapObject; // Aplicación SAP2000
        public cSapModel mySapModel; // Fichero de SAP dentro del programa
        string ProgramPath = @"C:\Program Files\Computers and Structures\SAP2000 25\SAP2000.exe"; // Asignamos la ruta de la aplicación SAP2000 para ejecutarlo
        eUnits UnidadesIniciales; // Establecer unidades iniciales en SAP2000



        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------

        // Escribir todos los métodos de clase aquí. Añadir un docstring para
        // tener información acerca de su funcionamiento y parámetros de entrada
        // y salida. Recordar añadirlo al excel de registro de métodos. Poner 
        // todos los métodos públicos para evitar errores de acceso.

        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------



        /// <summary>
        /// Busca todos los archivos de SAP (.sdb) en una carpeta a partir de una ruta 
        /// y te devuelve una lista con todas las rutas de los archivos SAP.
        /// </summary>
        /// <param name="SAPFolderRoute">
        /// Ruta de la carpeta donde buscar los archivos SAP (string). 
        /// </param>
        /// <returns>Lista de strings con las rutas de todos los archivos SAP en esa carpeta.</returns>
        public List<string> SearchSAPFiles(string SAPFolderRoute)
        {
            List<string> SAPFilesRoute = new List<string>();

            foreach (string file in Directory.GetFiles(SAPFolderRoute, "*.sdb", SearchOption.AllDirectories))
            {
                SAPFilesRoute.Add(file);
            }

            return SAPFilesRoute;
        }


        //---------------------------------------------------------------------------------


        /// <summary>
        /// Abre la aplicación SAP2000 y te devuelve la instancia del objeto.
        /// </summary>
        /// <returns>Instancia del objecto SAP.</returns>
        public cOAPI OpenSAPObject()
        {
            cHelper myHelper = new Helper();
            cOAPI mySapObject = null;

            myHelper = (cHelper)Activator.CreateInstance(Type.GetTypeFromProgID("SAP2000v1.Helper", true));
            mySapObject = myHelper.CreateObject(ProgramPath);
            mySapObject.ApplicationStart(eUnits.N_mm_C);

            return mySapObject;
        }


        //---------------------------------------------------------------------------------


        /// <summary>
        /// Inicializa un modelo de SAP2000 a partir de una instancia del objecto (SapObject).
        /// </summary>
        /// <param name="SapObject">
        /// Instancia del objecto SAP (SapObject). 
        /// </param>
        /// <returns>Instancia del modelo de SAP2000 (SapModel).</returns>
        public cSapModel OpenSAPModel(cOAPI SapObject)
        {
            mySapModel = SapObject.SapModel;
            mySapModel.InitializeNewModel();

            return mySapModel;
        }


        //---------------------------------------------------------------------------------


        /// <summary>
        /// Carga un archivo .sdb a partir de su ruta y de la instancia del modelo (SapModel).
        /// </summary>
        /// <param name="SapModel">
        /// Instancia del modelo SAP (SapModel). 
        /// </param>
        /// <param name="SapFileRoute">
        /// Ruta del fichero .sdb de SAP2000 que se desea cargar (string). 
        /// </param>
        public void LoadModels(cSapModel SapModel, string SAPFileRoute)
        {
            SapModel.File.OpenFile(SAPFileRoute);
        }


        //---------------------------------------------------------------------------------


        /// <summary>
        /// Calcula un archivo .sdb abierto a partir de la instancia del modelo (SapModel). 
        /// Es necesario que la instancia SapModel tenga cargado un fichero calculable.
        /// </summary>
        /// <param name="SapModel">
        /// Instancia del modelo SAP (SapModel) con un fichero calculable cargado. 
        /// </param>
        public void RunModel(cSapModel SapModel)
        {
            SapModel.Analyze.RunAnalysis();
        }


        //---------------------------------------------------------------------------------


        /// <summary>
        /// Selecciona las hipótesis que se quieren analizar para sacar algún output 
        /// (como reacciones o esfuerzos). Se seleccionan a partir de un string con 
        /// el nombre de la hipótesis deseada (por ejemplo "ULS") y de una instancia 
        /// del modelo (SapModel). Se debe incluir un bool que si es true deselecciona 
        /// todas las hipótesis, y si es false las deja como estaban. Se recomienda 
        /// poner true la primera vez que se use este método. Si se desean seleccionar 
        /// varias hipótesis, utilizar este método tantas veces como se requiera.
        /// </summary>
        /// <param name="SapModel">
        /// Instancia del modelo SAP (SapModel) con un fichero calculado cargado. 
        /// </param>
        /// <param name="Combo">
        /// Nombre de la hipótesis que se desea seleccionar (string). 
        /// </param>
        /// <param name="Deselect">
        /// Bool que si es true deselecciona todas las hipótesis seleccionadas. 
        /// </param>
        public void SelectHypotesis(cSapModel SapModel, string Combo, bool Deselect)
        {
            if (Deselect == true)
            {
                SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();
            }
            SapModel.Results.Setup.SetComboSelectedForOutput(Combo);
        }


        //---------------------------------------------------------------------------------


        /// <summary>
        /// Extrae las tablas que se quieran en excel a partir de un array con 
        /// el nombre de las tablas a extraer y de la instancia del modelo con 
        /// un fichero calculado cargado.
        /// </summary>
        /// <param name="SapModel">
        /// Instancia del modelo SAP (SapModel) con un fichero calculado cargado. 
        /// </param>
        /// <param name="TableKey">
        /// Array de strings con los nombres de las tablas a extraer. 
        /// </param>
        public void ExtractDataInExcel(cSapModel SapModel, string[] TableKey)
        {
            int WindowHandle = 1;
            SapModel.DatabaseTables.ShowTablesInExcel(ref TableKey, WindowHandle);
        }


        //---------------------------------------------------------------------------------


        /// <summary>
        /// Cierra la aplicación y limpia las instancias del SapModel y del 
        /// SapObject. Después de este método, si se quiere cargar otro fichero 
        /// se deberá volver a inicializar el SapObject y el SapModel.
        /// </summary>
        /// <param name="SapObject">
        /// Instancia del objecto SAP (SapObject). 
        /// </param>
        /// <param name="SapModel">
        /// Instancia del modelo SAP (SapModel). 
        /// </param>
        public void CloseModels(cOAPI SAPObject, cSapModel SapModel)
        {
            SAPObject.ApplicationExit(true);
            SAPObject = null;
            SapModel = null;

            GC.Collect(); // Forzar recolección de basura para limpiar instancias
            GC.WaitForPendingFinalizers();
        }


        //---------------------------------------------------------------------------------



    }
}

