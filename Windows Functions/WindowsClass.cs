using System;
using Microsoft.Win32;
using System.Windows;


namespace WindowsMethods
{
    public class WindowsClass
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
        /// Abre una ventana de Windows y te permite seleccionar un archivo con 
        /// cualquier extensión. Te devuelve un string con su ruta. Si no se selecciona 
        /// nada te devuelve un string vacío.
        /// </summary>
        /// <returns>
        /// Ruta del archivo seleccionado (string). Si no se selecciona nada te devuelve 
        /// un string vacío.
        /// </returns>
        public string SearchFile()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Seleccionar archivo",
                Filter = "Todos los archivos (*.*)|*.*",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };

            if (openFileDialog.ShowDialog() == true)
            {
                return openFileDialog.FileName;
            }

            return string.Empty;
        }


        //---------------------------------------------------------------------------------


        /// <summary>
        /// Abre una ventana de Windows y te permite seleccionar un archivo de SAP con 
        /// extensión ".sdb". Te devuelve un string con su ruta. Si no se selecciona 
        /// nada te devuelve un string vacío.
        /// </summary>
        /// <returns>
        /// Ruta del archivo seleccionado (string). Si no se selecciona nada te devuelve 
        /// un string vacío.
        /// </returns>
        public string SearchSAPFile()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Seleccionar archivo",
                Filter = "Archivos SDB (*.sdb)|*.sdb",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };

            if (openFileDialog.ShowDialog() == true)
            {
                return openFileDialog.FileName;
            }

            return string.Empty;
        }


        //---------------------------------------------------------------------------------


        /// <summary>
        /// Abre una ventana de Windows y te permite seleccionar una carpeta. Te 
        /// devuelve un string con su ruta. Si no se selecciona nada te devuelve 
        /// un string vacío.
        /// </summary>
        /// <returns>
        /// Ruta de la carpeta seleccionada (string). Si no se selecciona nada te 
        /// devuelve un string vacío.
        /// </returns>
        public string SearchFolder()
        {
            OpenFolderDialog openFolderDialog = new OpenFolderDialog
            {
                Title = "Seleccionar archivo",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };

            if (openFolderDialog.ShowDialog() == true)
            {
                return openFolderDialog.FolderName;
            }

            return string.Empty;
        }


        //---------------------------------------------------------------------------------


        /// <summary>
        /// Abre ventanas para que selecciones los archivos de SAP de posición de 
        /// defensa, intermedia y de resguardo. Añade sus rutas a un array de strings 
        /// y los pone en el orden mencionado antes. Se debe especificar cuál es el 
        /// índice de la posición del array donde colocar estas rutas de archivos 
        /// (el índice especificado será la ruta de la posición de defensa, el siguiente 
        /// será la posición intermedia, y la siguiente la de resguardo). Si no se 
        /// selecciona alguno de los tres archivos, se añadirá una ruta vacía al array.
        /// </summary>
        /// <param name="FileRouteList">
        /// Array de strings donde guardar las rutas de los archivos SAP. Debe ser de 
        /// tamaño mínimo 3 para poder albergar estas tres rutas de archivos. 
        /// </param>
        /// <param name="index">
        /// Índice de la posición del array en el que guardar la primera de las tres 
        /// rutas de archivos SAP. Las dos siguients rutas se guardarán en los índices 
        /// sucesivos. 
        /// </param>
        public void StoreFileRoutes(string[] FileRouteList, int index)
        {
            MessageBox.Show("Selecciona el archivo de posicion de defensa");
            FileRouteList[index] = SearchSAPFile();

            MessageBox.Show("Selecciona el archivo de posicion intermedia");
            FileRouteList[index+1] = SearchSAPFile();

            MessageBox.Show("Selecciona el archivo de posicion de funcionamiento");
            FileRouteList[index+2] = SearchSAPFile();

        }


        //---------------------------------------------------------------------------------



    }
}