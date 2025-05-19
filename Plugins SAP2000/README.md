# Plantilla desarrollo de plugins en entorno SAP2000
Repositorio que contiene un proyecto basico para crear plugings de SAP2000.

El proyecto contiene lo basico para crear un pluging que muestre una ventana y un boton ya vinculado.
Tambien contiene la funcion y la clase que realizan la conexion con SAP2000.

Para usar este archivo base sera necesario copiar las clases que contiene el namespace del archivo ejemplo
en nuestro proyecto creado en visual studio. La plantilla que se debe usar en visual studio es:
"Biblioteca de clases (.NET Framework) Proyecto para crear una biblioteca de calses de C# (.dll)."

Para compilar el archivo usaremos la combinacion de Botones: "Ctrl + B". En caso de que no funcion el atajo,
en la ventan de herramientas iremos a:

    Compilar -> Generar (Nombre de tu archivo)

Para vincular el plugin con SAP200, debemos:

    1.- Copiar nuestro archivo "(Nombre del proyecto).dll" a el directorio que queramos (Funciona subiendo el
    archivo al sharepoint y vinculandolo directamente). El archivo .dll se encuentra en la carpeta del proyecto:

    "Carpeta de proyecto" -> bin -> Debug

    2.- Vincularemos el archivo .dll en SAP200 desde la ventana de herramientas: 

    Tools -> Add/Show plugins... -> Browse (Buscamos la ruta del archivo .dll) -> Cambiamos el Nombre y Texto del menu -> Add

Para ejecutar el plugin:

    Ventana de herramientas -> Tools -> "Nombre Plugin"

# SOlUCION DE ERRORES
    -Despues de ejecutar el Plugin SAP2000 no responde: Hay que asegurarse que en nuestra funcion main, al final del "try", hayamos 
    añadido la funcion "ISapPlugin.Finish(0);"


