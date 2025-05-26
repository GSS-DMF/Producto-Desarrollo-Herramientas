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
            Design = new DesignSubclass(this);
            ElementFinder = new ElementFinderSubclass(this);
        }

        public FileManagerSubclass FileManager { get; }

        public AnalysisSubclass Analysis { get; }

        public ExcelTablesSubclass ExcelTables { get; }

        public DesignSubclass Design { get; }

        public ElementFinderSubclass ElementFinder { get; }

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
            /// Abre SAP en segundo plano
            /// </summary>
            /// <returns>
            /// Devuelve el objeto SAP abierto
            /// </returns>
      
            public cOAPI OpenSAPObjectHidden()
            {
                cHelper myHelper = new Helper();
                cOAPI mySapObject = null;

                myHelper = (cHelper)Activator.CreateInstance(Type.GetTypeFromProgID("SAP2000v1.Helper", true));
                mySapObject = myHelper.CreateObject(ProgramPath);
                mySapObject.ApplicationStart(eUnits.N_mm_C,false);

                return mySapObject;
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


            /// <summary>
            /// Obtiene una lista con los nombres de los elementos de un modelo según la tipología que se elija
            /// </summary>
            /// <param name="objectType">
            /// Tipología de elemento a seleccionar
            /// 1:Point Object
            /// 2:Frame Object
            /// 3:Cable Object
            /// 4:Tendom Object
            /// 5:Area Object
            /// 6:Solid Object
            /// 7:Link Object
            /// </param>
            /// <returns>
            /// Devuelve una lista cn los nombres de los elementos del modelo
            /// </returns>
            public string[] GetElements(cSapModel mySapModel, int objectType)
            {
                //Seleccionamos todo en el modelo
                mySapModel.SelectObj.All();

                int NumberItems = 0;
                int[] Type = new int[1];
                string[] ObjectName = new string[1];

                mySapModel.SelectObj.GetSelected(ref NumberItems, ref Type, ref ObjectName);

                //Obtenemos el listado de elementos en función de lo que se necesite: 

                List<string> nombres = new List<string>();
                for (int i = 0; i < NumberItems; i++)
                {
                    if (Type[i] == objectType)
                    {
                        nombres.Add(ObjectName[i]);
                    }
                }

                string[] elementos = nombres.ToArray();

                mySapModel.SelectObj.ClearSelection();

                return elementos;
            }

            //---------------------------------------------------------------------------------


            /// <summary>
            /// Dado el nombre de una barra de un modelo SAP, devuelve como string la altura,
            /// el espesor, y el material del perfil SHS asignado a la barra como un double[] {"B","e","fy"}
            /// </summary>
            /// <param name="mySapModel">
            /// Instancia del modelo SAP2000
            /// </param>
            /// <param name="frameName">
            /// Nombre de la barra
            /// </param>
            /// <returns>
            /// altura B y el espesor e del perfil SHS asignado a la barra como un double[] {"B","e","fy"}
            /// </returns>
            public static double[] GetSHSProperties(cSapModel mySapModel, string frameName)
            {
                string PropName = "";
                string SAuto = "";

                mySapModel.FrameObj.GetSection(frameName, ref PropName, ref SAuto);

                string[] partes = PropName.Split('/');
                string[] perfil = partes[0].Split("-");
                string[] dimensiones = perfil[1].Split("x");
                int r = 7;
                dimensiones[0] = dimensiones[0].Trim();
                dimensiones[1] = dimensiones[1].Trim();
                double.TryParse(dimensiones[0], out double B);
                double.TryParse(dimensiones[1], out double e);

                Match valor = Regex.Match(partes[1], @"\d+");
                double.TryParse(valor.Value, out double fy);

                return new double[] { B, e, fy };
            }

            //---------------------------------------------------------------------------------

            /// <summary>
            /// Devuelve la envolvente de esfuerzos de unas barras en un punto determinado 
            /// El resultado es un array con los máximos esfuerzos de todo el conjunto de barras
            /// {N=P, Vy=V2, Vz=V3, Mt=T, My=M2, Mz=M3} Los esfuerzos máximos no tinen por qué darse
            /// todos en la misma barra del conjunto
            /// </summary>
            /// <param name="mySapModel">
            /// Objeto SAP2000
            /// </param>
            /// <param name="combo">
            /// Nombre de la combinación en la que se necistan los esfuerzos
            /// </param>
            /// <param name="frames">
            /// Array con los nombres de las barras a evaluar
            /// </param>
            /// <param name="point">
            /// Posición del punto a analizar (0-L)
            /// </param>
            /// <returns>
            /// Devuelve un array con la envolvente de esfuerzos del conjunto de barras
            /// </returns>
            public static double[] GetFrameForces(cSapModel mySapModel,string combo, string[] frames, double point)
            {
                //Cambiar unidades, seleccionar hipótesis y analizar el modelo
                mySapModel.SetPresentUnits(eUnits.kN_m_C);
                SAP.AnalysisSubclass.RunModel(mySapModel);
                SAP.AnalysisSubclass.SelectHypotesis(mySapModel, combo, true);

                // Inicializar arrays de resultados globales
                double[] N = new double[frames.Length];
                double[] Vy = new double[frames.Length];
                double[] Vz = new double[frames.Length];
                double[] Mt = new double[frames.Length];
                double[] My = new double[frames.Length];
                double[] Mz = new double[frames.Length];

                //Variables de salida inicializadas para SAP2000
                int NumberResults = 5000;
                string[] Obj = new string[1];
                double[] ObjSta = new double[1];
                string[] Elm = new string[1];
                double[] ElmSta = new double[1];
                string[] LoadCase = new string[1];
                string[] StepType = new string[1];
                double[] StepNum = new double[1];
                double[] P = new double[1];
                double[] V2 = new double[1];
                double[] V3 = new double[1];
                double[] T = new double[1];
                double[] M2 = new double[1];
                double[] M3 = new double[1];

                for (int i = 0; i < frames.Length; i++)
                {
                    // Seleccionar el marco actual
                    mySapModel.FrameObj.SetSelected(frames[i], true, eItemType.Objects);

                    // Obtener resultados de esfuerzos
                    mySapModel.Results.FrameForce(frames[i], eItemTypeElm.ObjectElm, ref NumberResults, ref Obj, ref ObjSta, ref Elm, ref ElmSta, ref LoadCase, ref StepType, ref StepNum, ref P, ref V2, ref V3, ref T, ref M2, ref M3);

                    // Filtrar los esfuerzos en el punto deseado
                    var esfuerzos = Enumerable.Range(0, ObjSta.Length)
                        .Where(j => ObjSta[j] == point)
                        .Select(j => new
                        {
                            N = Math.Abs(P[j]),
                            Vy = Math.Abs(V2[j]),
                            Vz = Math.Abs(V3[j]),
                            Mt = Math.Abs(T[j]),
                            My = Math.Abs(M2[j]),
                            Mz = Math.Abs(M3[j])
                        }).ToList();

                    // Asignar el máximo de cada esfuerzo al arreglo correspondiente
                    if (esfuerzos.Any())
                    {
                        N[i] = esfuerzos.Max(e => e.N);
                        Vy[i] = esfuerzos.Max(e => e.Vy);
                        Vz[i] = esfuerzos.Max(e => e.Vz);
                        Mt[i] = esfuerzos.Max(e => e.Mt);
                        My[i] = esfuerzos.Max(e => e.My);
                        Mz[i] = esfuerzos.Max(e => e.Mz);
                    }
                }

                return new double[] { N.Max(), Vy.Max(), Vz.Max(), Mt.Max(), My.Max(), Mz.Max() };
            }
        
            //---------------------------------------------------------------------------------

            /// <summary>
            /// Devuelve la envolvente de esfuerzos de una barra 
            /// El resultado es un array con los máximos esfuerzos de la barra
            /// {N=P, Vy=V2, Vz=V3, Mt=T, My=M2, Mz=M3} 
            /// </summary>
            /// <param name="mySapModel">
            /// Objeto SAP2000
            /// <param name="frame">
            /// Array con los nombres de las barras a evaluar
            /// <returns>
            /// Devuelve un array con la envolvente de esfuerzos de la barra
            /// </returns>
            public static double[] GetOneFrameForces(cSapModel mySapModel,string combo, string frame)
            {
                mySapModel.SetPresentUnits(eUnits.kN_m_C);

                int NumberResults = 5000;
                string[] Obj = new string[1], Elm = new string[1], LoadCase = new string[1], StepType = new string[1];
                double[] ObjSta = new double[1], ElmSta = new double[1], StepNum = new double[1], P = new double[1], V2 = new double[1], V3 = new double[1], T = new double[1], M2 = new double[1], M3 = new double[1];

                RunModel(mySapModel);
                SelectHypotesis(mySapModel, combo, true);

                int ret = mySapModel.FrameObj.SetSelected(frame, true, eItemType.Objects);
                ret= mySapModel.Results.FrameForce(frame, eItemTypeElm.ObjectElm, ref NumberResults, ref Obj, ref ObjSta, ref Elm, ref ElmSta, ref LoadCase, ref StepType, ref StepNum, ref P, ref V2, ref V3, ref T, ref M2, ref M3);

                double N = Math.Max(Math.Abs(P.Max()), Math.Abs(P.Min()));
                double Vy = Math.Max(Math.Abs(V2.Max()), Math.Abs(V2.Min()));
                double Vz = Math.Max(Math.Abs(V3.Max()), Math.Abs(V3.Min()));
                double Mt = Math.Max(Math.Abs(T.Max()), Math.Abs(T.Min()));
                double My = Math.Max(Math.Abs(M2.Max()), Math.Abs(M2.Min()));
                double Mz = Math.Max(Math.Abs(M3.Max()), Math.Abs(M3.Min()));

                return new double[] {N,Vy,Vz,Mt,My,Mz };
            }


            //---------------------------------------------------------------------------------




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
    
        public class DesignSubclass // Clase par las funciones de apoyo para dimensionar perfiles
        {
            private readonly SAP _sap;

            public DesignSubclass(SAP sap)
            {
                _sap = sap;
            }
            
            //---------------------------------------------------------------------------------

            /// <summary>
            /// Busca el nombre del perfil (previamente seleccionado en el modelo) en la listaperfiles
            /// y lo cambia por el siguiente de la lista
            /// </summary>
            /// <param name="mySapModel">
            /// Instancia del modelo SAP
            /// </param>
            /// <param name="listaperfiles">
            /// Lista con los perfiles ordenados según se quiera
            /// </param>
            /// <returns>
            /// Devuelve el nombre del siguiente perfil de la lista.
            /// </returns>
            public static string ChangeSection(cSapModel mySapModel, string[] listaperfiles)
            {
                int ret = 0;
                int numberItems = 0;
                int[] objectType = new int[1];
                string[] itemName = new string[1];
                mySapModel.SelectObj.GetSelected(ref numberItems, ref objectType, ref itemName);

                string Section = "";
                ret = mySapModel.DesignColdFormed.GetDesignSection(itemName[0], ref Section);
                if (Section == "")
                {
                    ret = mySapModel.DesignSteel.GetDesignSection(itemName[0], ref Section);
                }
                int Pos = 0;
                foreach (string perfil in listaperfiles)
                {
                    if (Section == perfil)
                    {
                        foreach (var item in itemName)
                        {
                            mySapModel.FrameObj.SetSection(item, listaperfiles[Pos + 1]);
                        }
                        break;
                    }

                    else
                    {
                        Pos = Pos + 1;
                    }
                }

                return "";
            }

            //---------------------------------------------------------------------------------

            /// <summary>
            /// Calcula el desplazamiento de un nudo, en SLS
            /// </summary>
            /// <param name="mySapModel">
            /// Instancia de SAP2000
            /// </param>
            /// <param name="joint">
            /// Nudo a evaluar
            /// </param>
            /// <returns></returns>
            public static double JointDisplacement(cSapModel mySapModel, string joint)
            {
                int NumberResults = 0;
                string[] Obj = new string[1], Elm = new string[1], LoadCase = new string[1], StepType = new string[1];
                double[] Stepnum = new double[1], U1 = new double[1], U2 = new double[1], U3 = new double[1], R1 = new double[1], R2 = new double[1], R3 = new double[1];

                SAP.AnalysisSubclass.SelectHypotesis(mySapModel, "SLS", true);

                int ret=mySapModel.Results.JointDispl(joint,eItemTypeElm.ObjectElm,ref NumberResults,ref Obj,ref Elm,ref LoadCase,ref StepType,ref Stepnum,ref U1,ref U2,ref U3,ref R1,ref R2,ref R3);

                double maxUx = Math.Max(Math.Abs(U1[0]), Math.Abs(U1[1]));
                double maxUy = Math.Max(Math.Abs(U2[0]), Math.Abs(U2[1]));
                double maxUz = Math.Max(Math.Abs(U3[0]), Math.Abs(U3[1]));

                double d = 0;

                if(ret==0)
                {
                    d = Math.Sqrt(maxUx * maxUx + maxUy * maxUy + maxUz * maxUz);
                }

                return d;
            }

            //---------------------------------------------------------------------------------

            /// <summary>
            /// Hace la comprobación de Torsor/cortante que no hace SAP2000
            /// </summary>
            /// <param name="mySapModel">
            /// Instancia de SAP2000
            /// </param>
            /// <param name="barra">
            /// Barra que se quiere analizar
            /// </param>
            /// <param name="punto">
            /// Posición de la barra que se quiere analizar
            /// </param>
            /// <returns>
            /// Devuelve el aprovechamiento en cortante y torsión combinadas
            /// </returns>
            public static double[] ShearTorsionInteractionCheck(cSapModel mySapModel, string barra, double punto)
            {
                //Sacamos del modelo los datos de diseño necesarios. Unidades en N y mm
                mySapModel.SetPresentUnits(eUnits.N_mm_C);
                SAP.AnalysisSubclass.SelectHypotesis(mySapModel, "ULS", true);

                string PropName = "", SAuto = "";
                double[] prop=SAP.AnalysisSubclass.GetSHSProperties(mySapModel, barra);

                double gamma = 0;
                mySapModel.DesignColdFormed.EuroCold06.GetPreference(8, ref gamma);

                //Obtenemos esfuerzos. Unidades en kN y m
                mySapModel.SetPresentUnits(eUnits.kN_m_C);
                SAP.AnalysisSubclass.RunModel(mySapModel);

                double[] esfuerzos = SAP.AnalysisSubclass.GetFrameForces(mySapModel,"ULS", new[] {barra}, punto);

                double VcEd = Math.Max(esfuerzos[1], esfuerzos[2]);
                double MtEd = esfuerzos[3];
                    
                //Formulación
                double d = prop[0] - (2 * prop[1]) - (2 * 7);
                double Av = 2 * (prop[0] - (2 * prop[1]));
                double fyd = prop[2] / gamma;
                double VplRd = (Av * fyd) / (Math.Sqrt(3) * 1000);
                double Wt = (2 * prop[1] * Math.Pow(prop[0] - prop[1], 2)) / 1000;
                double TaoTEd = (MtEd * 1000) / Wt;
                double VplTEd = VplRd * (1 - (TaoTEd / fyd / Math.Sqrt(3)));
                double MtRd = (1 / Math.Sqrt(3)) * Wt * fyd / 1000;
                    
                //Ratios
                double AprV = VcEd / VplTEd * 100;
                double AprM = MtEd / MtRd * 100;
            
                return new[] {Math.Round(AprV,0), Math.Round(AprM,0)}; 
            }
        
            //---------------------------------------------------------------------------------

            /// <summary>
            /// Calcula el vano (distancia libre) entre dos pilares consecutivos que rodean a un nudo específico,
            /// utilizando sus coordenadas en el eje Y dentro del modelo SAP2000.
            /// </summary>
            /// <param name="mySapModel">
            /// Instancia del modelo SAP2000 (cSapModel) desde la cual se obtienen las coordenadas.
            /// </param>
            /// <param name="joint">
            /// Nombre del nudo (joint) para el cual se desea calcular el vano.
            /// </param>
            /// <param name="piles">
            /// Array de identificadores de los pilares (joints) que definen los extremos del vano.
            /// </param>
            /// <returns>
            /// Distancia en el eje Y entre los dos pilares más cercanos que rodean al nudo especificado. 
            /// Si no se encuentra un vano válido, devuelve 0.
            /// </returns>
            public static double FindSpan(cSapModel mySapModel, string joint, string[] piles)
            {
                double X = 0, Y=0, Z = 0;

                mySapModel.PointElm.GetCoordCartesian(joint, ref X, ref Y, ref Z);
                double coordNudoY = Y;

                double[] coordPilaresY= new double[piles.Length];
                for(int i=0;i<piles.Length;i++)
                {
                    double px = 0, py = 0, pz = 0;
                    mySapModel.PointElm.GetCoordCartesian(piles[i], ref px, ref py, ref pz);
                    coordPilaresY[i] = py;
                }

                // Ordenar coordenadas
                Array.Sort(coordPilaresY);

                // Buscar vano: distancia entre los dos pilares más cercanos que rodean al nudo
                double vano = 0;
                for (int i = 0; i < coordPilaresY.Length - 1; i++)
                {
                    if (coordPilaresY[i] <= coordNudoY && coordNudoY <= coordPilaresY[i + 1])
                    {
                        vano = coordPilaresY[i + 1] - coordPilaresY[i];
                        break;
                    }
                }

                return vano;
            }

            //---------------------------------------------------------------------------------



            //---------------------------------------------------------------------------------



            //---------------------------------------------------------------------------------



            //---------------------------------------------------------------------------------
        }
    
        public class ElementFinderSubclass // Clase para las funciones que devuelven nombres de barras y nudos
        {
            private readonly SAP _sap;

            public TrackerSubclass _tracker { get; }
            public FixedSubclass _fixed { get; }

            public ElementFinderSubclass(SAP sap)
            {
                _sap = sap;
                _tracker=new TrackerSubclass(this);
                _fixed = new FixedSubclass(this);
            }

            public class TrackerSubclass // Funciones para trackers
            {
                private readonly ElementFinderSubclass _elementFinder;

                public TrackerSubclass(ElementFinderSubclass elementFinder)
                {
                    _elementFinder = elementFinder;
                }

                //---------------------------------------------------------------------------------

                /// <summary>
                /// Cuenta el número de vigas específicas en el modelo SAP2000 seleccionando ciertos objetos de marco.
                /// </summary>
                /// <param name="mySapModel">
                /// Instancia del modelo SAP2000 
                /// </param>
                /// <returns>
                /// El número total de vigas encontradas y seleccionadas correctamente
                /// </returns>
                public static int BeamNumber(cSapModel mySapModel)
                {
                    mySapModel.SelectObj.ClearSelection();

                    int nvigas = 0;

                    int ret = mySapModel.FrameObj.SetSelected("B1", true, eItemType.Objects);
                    if (ret == 0) { nvigas++; }

                    ret = mySapModel.FrameObj.SetSelected("B1_Motor", true, eItemType.Objects);
                    if (ret == 0) { nvigas++; }

                    mySapModel.SelectObj.ClearSelection();

                    for (int i = 2; i <= 6; i++)
                    {
                        string viga = "B" + i;
                        ret = mySapModel.FrameObj.SetSelected(viga, true, eItemType.Objects);
                        if (ret == 0)
                        {
                            nvigas++;
                        }
                    }
                    mySapModel.SelectObj.ClearSelection();

                    return nvigas;
                }


                //---------------------------------------------------------------------------------

                /// <summary>
                /// Devuelve un array con los nombres de las vigas del lado norte que existen en el modelo SAP2000.
                /// </summary>
                /// <param name="mySapModel">
                /// Instancia del modelo SAP2000 
                /// </param>
                /// <returns>
                /// Un array de cadenas con los nombres de las vigas encontradas
                /// </returns>
                public static string[] NorthBeams(cSapModel mySapModel)
                {
                    int nvigas = BeamNumber(mySapModel);
                    int contador = 0;

                    string[] vigas = new string[nvigas];

                    int ret = mySapModel.FrameObj.SetSelected("B1", true, eItemType.Objects);
                    if (ret == 0) 
                    { 
                        vigas[contador++]="B1"; 
                    }

                    ret = mySapModel.FrameObj.SetSelected("B1_Motor", true, eItemType.Objects);
                    if (ret == 0)
                    {
                        vigas[contador++] = "B1_Motor";
                    }

                    for (int i = 1; i < nvigas; i++)
                    {
                        vigas[contador++] = "B" + i;
                    }

                    return vigas;
                }

                //---------------------------------------------------------------------------------

                /// <summary>
                /// Devuelve un array con los nombres de las vigas del lado sur que existen en el modelo SAP2000.
                /// </summary>
                /// <param name="mySapModel">
                /// Instancia del modelo SAP2000 
                /// </param>
                /// <returns>
                /// Un array de cadenas con los nombres de las vigas encontradas
                /// </returns>
                public static string[] SouthBeams(cSapModel mySapModel)
                {
                    int nvigas = BeamNumber(mySapModel);
                    int contador = 0;

                    string[] vigas = new string[nvigas];

                    int ret = mySapModel.FrameObj.SetSelected("B-1", true, eItemType.Objects);
                    if (ret == 0)
                    {
                        vigas[contador++] = "B-1";
                    }

                    ret = mySapModel.FrameObj.SetSelected("B-1_Motor", true, eItemType.Objects);
                    if (ret == 0)
                    {
                        vigas[contador++] = "B-1_Motor";
                    }

                    for (int i = 1; i < nvigas; i++)
                    {
                        vigas[contador++] = "B-" + i;
                    }

                    return vigas;
                }

                //---------------------------------------------------------------------------------

                /// <summary>
                /// Genera un array con los nombres de las uniones entre vigas del lado norte en formato "B1", "B2", ..., hasta "Bn".
                /// </summary>
                /// <param name="mySapModel">
                /// Instancia del modelo SAP2000 utilizada para determinar el número de vigas.
                /// </param>
                /// <returns>
                /// Un array de cadenas con los identificadores de las uniones entre vigas del lado norte.
                /// </returns>
                public static string[] NorthBC(cSapModel mySapModel)
                {
                    int nvigas = BeamNumber(mySapModel);

                    string[] BC_n = new string[nvigas + 1];

                    for (int i = 1; i <= nvigas; i++)
                    {
                        int j = i - 1;
                        BC_n[j] = "B" + i;
                    }
                    return BC_n;
                }

                //---------------------------------------------------------------------------------

                /// <summary>
                /// Genera un array con los nombres de las uniones entre vigas del lado sur en formato "B1", "B2", ..., hasta "Bn".
                /// </summary>
                /// <param name="mySapModel">
                /// Instancia del modelo SAP2000 utilizada para determinar el número de vigas.
                /// </param>
                /// <returns>
                /// Un array de cadenas con los identificadores de las uniones entre vigas del lado sur.
                /// </returns>
                public static string[] SouthBC(cSapModel mySapModel)
                {
                    int nvigas = BeamNumber(mySapModel);

                    string[] BC_n = new string[nvigas + 1];

                    for (int i = 1; i <= nvigas; i++)
                    {
                        int j = i - 1;
                        BC_n[j] = "B-" + i;
                    }
                    return BC_n;
                }

                //---------------------------------------------------------------------------------

                /// <summary>
                /// Cuenta el número de pilares específicas en el semitracker del modelo SAP2000. 
                /// Devuelve la mitad de los pilares generales más el motor
                /// </summary>
                /// <param name="mySapModel">
                /// Instancia del modelo SAP2000 
                /// </param>
                /// <returns>
                /// El número total de pilares encontradas y seleccionadas correctamente
                /// </returns>
                public static int PileNumber(cSapModel mySapModel)
                {
                    int npilares = 0;

                    for (int i = 0;i<=10;i++)
                    {
                        string pilar = "Column_" + i;
                        int ret = mySapModel.FrameObj.SetSelected(pilar, true, eItemType.Objects);
                        if (ret == 0)
                        {
                            npilares++;
                        }
                    }
                    return npilares;
                }

                //---------------------------------------------------------------------------------

                /// <summary>
                /// Devuelve un array con los nombres de los pilares del lado norte que existen en el modelo SAP2000.
                /// </summary>
                /// <param name="mySapModel">
                /// Instancia del modelo SAP2000 
                /// </param>
                /// <returns>
                /// Un array de cadenas con los nombres de los pilares encontrados
                /// </returns>
                public static string[] NorthPiles(cSapModel mySapModel)
                {
                    string[] pilares_n= new string[PileNumber(mySapModel)];

                    for (int i = 1;i<=PileNumber(mySapModel);i++)
                    {
                        int j=i - 1;
                        pilares_n[j] = "Column_" + i;
                    }

                    return pilares_n;
                }

                //---------------------------------------------------------------------------------

                /// <summary>
                /// Devuelve un array con los nombres de los pilares del lado sur que existen en el modelo SAP2000.
                /// </summary>
                /// <param name="mySapModel">
                /// Instancia del modelo SAP2000 
                /// </param>
                /// <returns>
                /// Un array de cadenas con los nombres de los pilares encontrados
                /// </returns>
                public static string[] SouthPiles(cSapModel mySapModel)
                {
                    string[] pilares_n = new string[PileNumber(mySapModel)];

                    for (int i = 1; i <= PileNumber(mySapModel); i++)
                    {
                        int j = i - 1;
                        pilares_n[j] = "Column_-" + i;
                    }

                    return pilares_n;
                }

                //---------------------------------------------------------------------------------

                /// <summary>
                /// Cuenta el número de vigas secundarias del norte en el modelo SAP2000. Si north=false
                /// devuelve el número de vigas al sur
                /// </summary>
                /// <param name="mySapModel">
                /// Instancia del modelo SAP2000
                /// </param>
                /// <param name="north">
                /// Si north=true (por defecto) devuelve el número de vigas al norte, si 
                /// north=false, devuelve el número de vigas al sur
                /// </param>
                /// <returns>
                /// El número total de vigas secundarias del lado norte o sur del tracker
                /// </returns>
                public static int SecundaryBeamNumber (cSapModel mySapModel, bool? north=true)
                {
                    int nsecundarias_n = 0;
                    string sb = "";

                    for (int i = 0; i <= 31; i++)
                    {
                        if(north==true)
                        { 
                            sb = "SBsN_" + i; 
                        }
                        else if(north==false)
                        {
                            sb = "SBsS_" + i;
                        }

                        int ret = mySapModel.FrameObj.SetSelected(sb, true, eItemType.Objects);

                        if (ret == 0)
                        {
                            nsecundarias_n++;
                        }
                    }
                    mySapModel.SelectObj.ClearSelection();

                    return nsecundarias_n;
                }

                //---------------------------------------------------------------------------------
                
                /// <summary>
                /// Devuelve los nombres de las secundarias al norte del tracker. Si sup=true (por defecto)
                /// devuelve el nombre de las vigas superiores, sino, el de las inferiores
                /// </summary>
                /// <param name="mySapModel">
                /// Instancia del modelo SAP2000
                /// </param>
                /// <param name="sup">
                /// Por defecto=true, vigas superiores, si false vigas inferiores
                /// </param>
                /// <returns>
                /// Devuelve los nombres de las secundarias al norte del tracker. Si sup=true (por defecto)
                /// devuelve el nombre de las vigas superiores, sino, el de las inferiores
                /// </returns>
                public static string[] NorthSecundaryBeams (cSapModel mySapModel, bool? sup=true)
                {
                    int nsecundarias = SecundaryBeamNumber(mySapModel, true);

                    string[] secundarias = new string[nsecundarias];

                    for(int i = 1;i<=nsecundarias;i++)
                    {
                        int j = i - 1;
                        if (sup == true)
                        {
                            secundarias[j] = "SBsN_" + i;
                        }
                        else if (sup == false)
                        {
                            secundarias[j] = "SBiN_" + i;
                        }
                    }

                    return secundarias;
                }
                
                //---------------------------------------------------------------------------------

                /// <summary>
                /// Devuelve los nombres de las secundarias al sur del tracker. Si sup=true (por defecto)
                /// devuelve el nombre de las vigas superiores, sino, el de las inferiores
                /// </summary>
                /// <param name="mySapModel">
                /// Instancia del modelo SAP2000
                /// </param>
                /// <param name="sup">
                /// Por defecto=true, vigas superiores, si false vigas inferiores
                /// </param>
                /// <returns>
                /// Devuelve los nombres de las secundarias al sur del tracker. Si sup=true (por defecto)
                /// devuelve el nombre de las vigas superiores, sino, el de las inferiores
                /// </returns>
                public static string[] SouthSecundaryBeams(cSapModel mySapModel, bool? sup = true)
                {
                    int nsecundarias = SecundaryBeamNumber(mySapModel, true);

                    string[] secundarias = new string[nsecundarias];

                    for (int i = 1; i <= nsecundarias; i++)
                    {
                        int j = i - 1;
                        if (sup == true)
                        {
                            secundarias[j] = "SBsS_" + i;
                        }
                        else if (sup == false)
                        {
                            secundarias[j] = "SBiS_" + i;
                        }
                    }

                    return secundarias;
                }

                //---------------------------------------------------------------------------------

                /// <summary>
                /// Devuelve los nombres de los refuerzos de las secundarias al norte del tracker. 
                /// Si sup=true (por defecto) devuelve el nombre de las vigas superiores, sino, 
                /// el de las inferiores
                /// </summary>
                /// <param name="mySapModel">
                /// Instancia del modelo SAP2000
                /// </param>
                /// <param name="sup">
                /// Por defecto=true,refuerzos de vigas superiores, si false de vigas inferiores
                /// </param>
                /// <returns>
                /// Devuelve los nombres de los refuerzos de las secundarias al norte del tracker. 
                /// Si sup=true (por defecto) devuelve el nombre de las vigas superiores, sino, 
                /// el de las inferiores
                /// </returns>
                public static string[] NorthSecundaryReinforcedBeams(cSapModel sapModel, bool? sup = true)
                {
                    int nsecundarias = SecundaryBeamNumber(mySapModel, true);

                    string[] secundarias = new string[nsecundarias];

                    for (int i = 1; i <= nsecundarias; i++)
                    {
                        int j = i - 1;
                        if (sup == true)
                        {
                            secundarias[j] = "SBsNr_" + i;
                        }
                        else if (sup == false)
                        {
                            secundarias[j] = "SBiNr_" + i;
                        }
                    }

                    return secundarias;
                }

                //---------------------------------------------------------------------------------

                /// <summary>
                /// Devuelve los nombres de los refuerzos de las secundarias al sur del tracker. 
                /// Si sup=true (por defecto) devuelve el nombre de las vigas superiores, sino, 
                /// el de las inferiores
                /// </summary>
                /// <param name="mySapModel">
                /// Instancia del modelo SAP2000
                /// </param>
                /// <param name="sup">
                /// Por defecto=true,refuerzos de vigas superiores, si false de vigas inferiores
                /// </param>
                /// <returns>
                /// Devuelve los nombres de los refuerzos de las secundarias al sur del tracker. 
                /// Si sup=true (por defecto) devuelve el nombre de las vigas superiores, sino, 
                /// el de las inferiores
                /// </returns>
                public static string[] SouthSecundaryReinforcedBeams(cSapModel sapModel, bool? sup = true)
                {
                    int nsecundarias = SecundaryBeamNumber(mySapModel, true);

                    string[] secundarias = new string[nsecundarias];

                    for (int i = 1; i <= nsecundarias; i++)
                    {
                        int j = i - 1;
                        if (sup == true)
                        {
                            secundarias[j] = "SBsSr_" + i;
                        }
                        else if (sup == false)
                        {
                            secundarias[j] = "SBiSr_" + i;
                        }
                    }

                    return secundarias;
                }

                //---------------------------------------------------------------------------------

                /// <summary>
                /// Obtiene los nodos iniciales o finales de un conjunto de barras de un modelo SAP2000
                /// </summary>
                /// <param name="mySapModel">
                /// Instancia del modelo SAP2000
                /// </param>
                /// <param name="frames">
                /// Array de nombres de barras de SAP2000
                /// </param>
                /// <param name="joint">
                /// Indicador del nodo a devolver:
                /// 1 para el nodo inicial (extremo i)
                /// 2 para el nodo final (extremo j)
                /// </param>
                /// <returns>
                /// Array con los nombres de los nudos correspondientes al extremo especificado
                /// </returns>
                public static string[] GetJoints(cSapModel mySapModel, string[]frames, int joint)
                {
                    int nbarras=frames.Length;

                    string[] point1= new string[nbarras];
                    string[] point2= new string[nbarras];

                    for (int i = 0;i< nbarras;i++)
                    {
                        mySapModel.FrameObj.GetPoints(frames[i],ref point1[i],ref point2[i]);
                    }

                    if(joint==1)
                    {
                        return point1;
                    }
                    else if(joint==2)
                    {
                        return point2;
                    }
                    else
                    {
                        return null;
                    }
                }

                //---------------------------------------------------------------------------------



                //---------------------------------------------------------------------------------



                //---------------------------------------------------------------------------------



                //---------------------------------------------------------------------------------



                //---------------------------------------------------------------------------------



                //---------------------------------------------------------------------------------
            }

            public class FixedSubclass // Funciones para trackers
            {
                private readonly ElementFinderSubclass _elementFinder;

                public FixedSubclass(ElementFinderSubclass elementFinder)
                {
                    _elementFinder = elementFinder;
                }

                //---------------------------------------------------------------------------------



                //---------------------------------------------------------------------------------



                //---------------------------------------------------------------------------------



                //---------------------------------------------------------------------------------



                //---------------------------------------------------------------------------------

            }
        }
    }
}