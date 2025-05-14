using SAP2000v1;
using System.IO;


namespace RepositorioFuncionesGitHub
{
    public class SAP
    {
        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------

        // Escribir todas las instancias de subclases aquí.

        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------



        public SAP()
        {
            // Constructor de la clase SAP
            FileManager = new FileManagerSubclass(this);
            Analysis = new AnalysisSubclass(this);
            ExcelTables = new ExcelTablesSubclass(this);
        }

        public FileManagerSubclass FileManager { get; }

        public AnalysisSubclass Analysis { get; }

        public ExcelTablesSubclass ExcelTables { get; }



        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------

        // Escribir todos los atributos de clase necesarios para los métodos aquí.
        // Añadir también una descripción de cada uno para poder localizarlos. Añadirlos 
        // todos como propiedades static para que las subclases tengan acceso a ellas.

        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------



        public static cOAPI mySapObject; // Aplicación SAP2000
        public static cSapModel mySapModel; // Fichero de SAP dentro del programa
        public static string ProgramPath = @"C:\Program Files\Computers and Structures\SAP2000 25\SAP2000.exe"; // Asignamos la ruta de la aplicación SAP2000 para ejecutarlo
        public static eUnits UnidadesIniciales; // Establecer unidades iniciales en SAP2000



        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------

        // Escribir todos los métodos de clase aquí. Añadir un docstring para
        // tener información acerca de su funcionamiento y parámetros de entrada
        // y salida. Recordar añadirlo al excel de registro de métodos. Poner 
        // todos los métodos públicos para evitar errores de acceso.

        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------







        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------

        // Escribir todas las subclases aquí. Añadir un docstring para tener 
        // información acerca de su funcionamiento y parámetros de entrada
        // y salida. Recordar añadirlo al excel de registro de métodos. Poner 
        // todos los métodos públicos para evitar errores de acceso.

        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------


        public class FileManagerSubclass // Clase para las funciones que gestionen ventanas y ficheros de SAP
        {
            //---------------------------------------------------------------------------------
            //---------------------------------------------------------------------------------

            // Traemos las propiedades de clase de la clase pricipal

            //---------------------------------------------------------------------------------
            //---------------------------------------------------------------------------------



            private readonly SAP _sap;

            public FileManagerSubclass(SAP sap)
            {
                _sap = sap;
            }



            //---------------------------------------------------------------------------------
            //---------------------------------------------------------------------------------

            // Introducimos los métodos de la subclase 

            //---------------------------------------------------------------------------------
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


        public class AnalysisSubclass // Clase para las funciones que hagan análisis (calcular, seleccionar hipótesis...)
        {
            //---------------------------------------------------------------------------------
            //---------------------------------------------------------------------------------

            // Traemos las propiedades de clase de la clase pricipal

            //---------------------------------------------------------------------------------
            //---------------------------------------------------------------------------------



            private readonly SAP _sap;

            public AnalysisSubclass(SAP sap)
            {
                _sap = sap;
            }



            //---------------------------------------------------------------------------------
            //---------------------------------------------------------------------------------

            // Introducimos los métodos de la subclase 

            //---------------------------------------------------------------------------------
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
            /// Calculamos la longitud de cualquier elemento de un modelo SAP2000, a partir del nombre de un segmento.
            /// </summary>
            /// <param name="SapModel">
            /// Instancia del modelo SAP (SapModel) con un fichero calculado cargado. 
            /// </param>
            /// <param name="elementName">
            /// Nombre del elemento del cual se quiere calcular la longitud.
            /// </param>
            public static double LongitudSegmento(cSapModel sapModel, string elementName)
            {
                double x1 = 0, y1 = 0, z1 = 0, x2 = 0, y2 = 0, z2 = 0;
                string point1 = "";
                string point2 = "";

                // Obtener las coordenadas de los nodos del elemento
                sapModel.FrameObj.GetPoints(elementName, ref point1, ref point2);
                sapModel.PointObj.GetCoordCartesian(point1, ref x1, ref y1, ref z1);
                sapModel.PointObj.GetCoordCartesian(point2, ref x2, ref y2, ref z2);

                // Calcular la longitud del elemento
                double length = Math.Sqrt(Math.Pow(x2 - x1, 2) + Math.Pow(y2 - y1, 2) + Math.Pow(z2 - z1, 2));

                return Math.Round(length, 2);
            }


            //---------------------------------------------------------------------------------



            /// <summary>
            /// Calculamos la longitud de cualquier elemento de refuerzo de un modelo SAP2000 2VR3, a partir del nombre de un refuerzo.
            /// Nombre de los refuerzos que se pueden calcular: SBsNr_x, SBiNr_x, SBsSr_x, SBiSr_x.
            /// Donde "x" es el numero de la viga secundaria del cual se quiere calcular la longitud.
            /// </summary>
            /// <param name="SapModel">
            /// Instancia del modelo SAP (SapModel) con un fichero calculado cargado. 
            /// </param>
            /// <param name="elementName">
            /// Nombre del elemento del cual se quiere calcular la longitud.
            /// </param>
            public static double LongitudRefuerzo(cSapModel sapModel, string elementName)
            {
                double x1 = 0, y1 = 0, z1 = 0, x2 = 0, y2 = 0, z2 = 0;
                string point1 = "";
                string point2 = "";

                elementName = elementName.Replace("_", "r_");

                // Obtener las coordenadas de los nodos del elemento
                sapModel.FrameObj.GetPoints(elementName, ref point1, ref point2);
                sapModel.PointObj.GetCoordCartesian(point1, ref x1, ref y1, ref z1);
                sapModel.PointObj.GetCoordCartesian(point2, ref x2, ref y2, ref z2);

                // Calcular la longitud del elemento
                double length = Math.Sqrt(Math.Pow(x2 - x1, 2) + Math.Pow(y2 - y1, 2) + Math.Pow(z2 - z1, 2));

                return Math.Round(length, 2);
            }


            //---------------------------------------------------------------------------------



            /// <summary>
            /// Calcula la longitud entre dos puntos. Partiendo de los nombres de los distintos puntos.
            /// </summary>
            /// <param name="SapModel">
            /// Instancia del modelo SAP (SapModel) con un fichero calculado cargado. 
            /// </param>
            /// <param name="point1">
            /// Nombre del primer punto del que se quiere calcular la distancia.
            /// </param>
            /// /// <param name="point2">
            /// Nombre del segundo punto del que se quiere calcular la distancia.
            /// </param>
            public static double LongitudEntrePuntos(cSapModel mySapModel, string point1, string point2)
            {
                int ret = 0;

                double[] coord_1 = new double[3];
                double[] coord_2 = new double[3];

                ret = mySapModel.PointObj.GetCoordCartesian(point1, ref coord_1[0], ref coord_1[1], ref coord_1[2]);
                ret = mySapModel.PointObj.GetCoordCartesian(point2, ref coord_2[0], ref coord_2[1], ref coord_2[2]);

                double Longitud = Math.Sqrt(Math.Pow(coord_1[0] - coord_2[0], 2) + Math.Pow(coord_1[1] - coord_2[1], 2) + Math.Pow(coord_1[2] - coord_2[2], 2));

                return Longitud;
            }


            //---------------------------------------------------------------------------------
        }

        public class ExcelTablesSubclass // Clase para las funciones que hagan análisis (calcular, seleccionar hipótesis...)
        {
            //---------------------------------------------------------------------------------
            //---------------------------------------------------------------------------------

            // Traemos las propiedades de clase de la clase pricipal

            //---------------------------------------------------------------------------------
            //---------------------------------------------------------------------------------



            private readonly SAP _sap;

            public ExcelTablesSubclass(SAP sap)
            {
                _sap = sap;
            }



            //---------------------------------------------------------------------------------
            //---------------------------------------------------------------------------------

            // Introducimos los métodos de la subclase 

            //---------------------------------------------------------------------------------
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
            /// Obtiene una tabla determinada de un modelo de SAP2000
            /// </summary>
            /// <param name="mySapModel">
            /// Modelo SAP del que obtener la tabla
            /// </param>
            /// <param name="tableName">
            /// nombre de la tabla a obtener
            /// </param>
            /// <returns>
            /// Devuelve la tabla completa del modelo SAP2000
            /// </returns>
            public string[,] GetTableArray(cSapModel mySapModel, string tableName)
            {
                int ret = 0;
                string[] FieldKeyList = new string[500];
                int TableVersion = 0;
                string[] FieldsKeysIncluded = new string[500];
                int NumberRecords = 0;
                string[] TableData = new string[500];

                ret = mySapModel.DatabaseTables.GetTableForDisplayArray(tableName, ref FieldKeyList, "All", ref TableVersion, ref FieldsKeysIncluded, ref NumberRecords, ref TableData);

                string[,] tabla = new string[NumberRecords + 1, FieldsKeysIncluded.Length];

                for (int i = 0; i < FieldsKeysIncluded.Length; i++)
                {
                    tabla[0, i] = FieldsKeysIncluded[i];
                }

                for (int i = 0; i < NumberRecords; i++)
                {
                    for (int j = 0; j < FieldsKeysIncluded.Length; j++)
                    {
                        tabla[i + 1, j] = TableData[i * FieldsKeysIncluded.Length + j];
                    }
                }

                return tabla;
            }
            

            //---------------------------------------------------------------------------------


            
        }
    }
}