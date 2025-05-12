# Producto-Desarrollo-Herramientas
Repositorio para guardar todas las funciones y código re-aprovechable en varios programas.

Se tiene el excel de registro de funciones añadidas, donde se incluirán el nombre de la función,
todos los inputs que sean necesarios (el tipo que sean como por ejemplo int o string, y la variable
en sí poniendo el nombre), los outputs que se den, y una descripción lo más detallada posible de 
lo que hace esa función.

Para poder usar las clases será necesario copiar el archivo en tu proyecto, instanciar la clase 
RepositorioFunciones e importar el namespace RepositorioFuncionesGitHub.

Pongamos un código de ejemplo para poder usarlo. Una vez copiamos la carpeta Class entera en nuestro
proyecto, vamos a nuestro fichero main y ponemos lo siguiente:


    // Importamos los namespaces
    using RepositorioFuncionesGitHub;


Una vez importado el namespace, lo siguiente es instanciar las clases en nuestro main. Al principio 
tendremos un código similar al siguiente:


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


Para instanciarlo, ponemos lo siguiente:

    namespace PROYECTO1
    {
        /// <summary>
        /// Interaction logic for MainWindow.xaml
        /// </summary>
        public partial class MainWindow : Window
        {
            // Instanciamos la clase
            RepositorioFunciones RepositorioFunciones = new RepositorioFunciones();

            .
            .
            .
        }
    }


Ahora ya podremos usar todos los métodos que queramos. Para hacerlo, se debe llamar a la instancia
de la clase en cuestión. Por ejemplo, si queremos usar el método RunModel de la clase SAP:


    RepositorioFunciones.SAP.RunModel(input1);


Si tenemos varias subclases dentro de una misma clase, seguimos la misma estructura (por ejemplo, 
la subclase Format dentro de la clase Tables):


    RepositorioFunciones.Tables.Format.Func1(input1);


De esta forma, se pueden usar todos los métodos incluidos en los archivos de las clases. Para saber
qué métodos están disponibles en esos archivos, consultar el excel de registro "DIRECTORIO_FUNCIONES".