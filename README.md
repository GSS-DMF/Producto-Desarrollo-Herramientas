# Producto-Desarrollo-Herramientas
Repositorio para guardar todas las funciones y código re-aprovechable en varios programas.

Se tiene el excel de registro de funciones añadidas, donde se incluirán el nombre de la función,
todos los inputs que sean necesarios (el tipo que sean como por ejemplo int o string, y la variable
en sí poniendo el nombre), los outputs que se den, y una descripción lo más detallada posible de 
lo que hace esa función.

Para poder usar las clases será necesario copiar el archivo en tu proyecto, instanciar la clase 
e importar estas clases.

Pongamos un código de ejemplo para poder usarlo. Una vez copiamos la carpeta Class entera en nuestro
proyecto, vamos a nuestro fichero main y ponemos lo siguiente:

    // Importamos los namespaces
    using SAPMethods;
    using WindowsMethods;
    using ExcelMethods;
    using WordMethods;
    using MathMethods;
    using TableMethods;

    // Importamos las librerias necesarias
    using SAP2000v1;
    using OfficeOpenXml;
    using System;
    using System.IO;
    using Microsoft.Win32;
    using System.Windows;
    using System.Runtime.InteropServices;
    using Excel = Microsoft.Office.Interop.Excel;
    using System.Runtime.Versioning;
    using System.Security;


Una vez importados los namespaces y librerias, lo siguiente es instanciar las clases en nuestro main. 
Al principio tendremos un código similar al siguiente:


    namespace PROYECTO1
    {
        /// <summary>
        /// Interaction logic for MainWindow.xaml
        /// </summary>
        public partial class MainWindow : Window
        {
            .
            .
            .
            .
        }
    }


Para instanciarlas, ponemos lo siguiente:

    namespace PROYECTO1
    {
        /// <summary>
        /// Interaction logic for MainWindow.xaml
        /// </summary>
        public partial class MainWindow : Window
        {
            // Instanciamos las clases
            SAPClass mySAPClass = new SAPClass();
            WindowsClass myWindowsClass = new WindowsClass();
            ExcelClass myExcelClass = new ExcelClass();
            WordClass myWordClass = new WordClass();
            MathClass myMathClass = new MathClass();

            .
            .
            .
        }
    }


Ahora ya podremos usar todos los métodos que queramos. Para hacerlo, se debe llamar a la instancia
de la clase en cuestión. Por ejemplo, si queremos usar el método RunModel de la clase SAPClass:


    mySAPClass.RunModel(input1);


De esta forma, se pueden usar todos los métodos incluidos en los archivos de las clases. Para saber
qué métodos están disponibles en esos archivos, consultar el excel de registro "DIRECTORIO_FUNCIONES".