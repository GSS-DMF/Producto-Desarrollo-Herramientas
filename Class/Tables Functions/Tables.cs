using System;
using Microsoft.Win32;
using System.Windows;


namespace RepositorioFuncionesGitHub
{
    public class Tables
    {
        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------

        // Escribir todas las instancias de subclases aquí. 

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







        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------

        // Escribir todas las subclases aquí. Añadir un docstring para tener 
        // información acerca de su funcionamiento y parámetros de entrada
        // y salida. Recordar añadirlo al excel de registro de métodos. Poner 
        // todos los métodos públicos para evitar errores de acceso.

        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------

        /// <summary>
        /// Mantiene las columnas de la tabla seleccionadas y elimina el resto
        /// </summary>
        /// <param name="table">
        /// Tabla original
        /// </param>
        /// <param name="columnNames">
        /// Nombres de las columnas que se quieren mantener en la tabla
        /// </param>
        /// <returns>
        /// Devuelve la tabla con las columnas seleccionadas
        /// </returns>
        public string[,] GetTableColumns(string[,] table, string[] columnNames)
        {
            int filas = table.GetLength(0);
            int columnas = columnNames.Length;
            string[,] nuevaTabla = new string[filas, columnas];

            //Encuentra los índices de las columnas deseadas
            int[] indicesColumnas = new int[columnas];
            for (int i = 0; i < columnas; i++)
            {
                for (int j = 0; j < table.GetLength(1); j++)
                {
                    if (table[0, j] == columnNames[i])
                    {
                        indicesColumnas[i] = j;
                        break;
                    }
                }
            }

            //Copia las columnas deseadas a la nueva tabla
            for (int i = 0; i < filas; i++)
            {
                for (int j = 0; j < columnas; j++)
                {
                    nuevaTabla[i, j] = table[i, indicesColumnas[j]];
                }
            }
            return nuevaTabla;
        }

        /// <summary>
        /// Filtra una tabla dada según si los valores de una columna coinciden con un valor dado 
        /// </summary>
        /// <param name="table">
        /// Tabla a filtrar
        /// </param>
        /// <param name="column">
        /// Columna para filtrar la tabla
        /// </param>
        /// <param name="value">
        /// Valor por el que se quiere filtrar la tabla
        /// </param>
        /// <returns>
        /// Devuelve la tabla filtrada
        /// </returns>
        public string[,] FilterTableEqual(string[,] table, string column, string value)
        {
            int filas = table.GetLength(0);
            int columnas = table.GetLength(1);
            int indiceColumna = -1;

            // Encontrar el índice de la columna
            for (int j = 0; j < columnas; j++)
            {
                if (table[0, j] == column)
                {
                    indiceColumna = j;
                    break;
                }
            }

            if (indiceColumna == -1)
            {
                throw new ArgumentException("Columna no encontrada");
            }

            // Crear una lista para almacenar las filas filtradas
            List<string[]> filasFiltradas = new List<string[]>();

            // Añadir la fila de encabezado
            filasFiltradas.Add(new string[columnas]);
            for (int j = 0; j < columnas; j++)
            {
                filasFiltradas[0][j] = table[0, j];
            }

            // Filtrar las filas según el criterio
            for (int i = 1; i < filas; i++)
            {
                bool agregarFila = false;
                string valor = table[i, indiceColumna];

                if (valor == value)
                {
                    agregarFila = true;
                }

                if (agregarFila)
                {
                    string[] fila = new string[columnas];
                    for (int j = 0; j < columnas; j++)
                    {
                        fila[j] = table[i, j];
                    }
                    filasFiltradas.Add(fila);
                }
            }

            // Convertir la lista de filas filtradas a una matriz bidimensional
            string[,] tablaFiltrada = new string[filasFiltradas.Count, columnas];
            for (int i = 0; i < filasFiltradas.Count; i++)
            {
                for (int j = 0; j < columnas; j++)
                {
                    tablaFiltrada[i, j] = filasFiltradas[i][j];
                }
            }

            return tablaFiltrada;
        }

        /// <summary>
        /// Filtra una tabla dada según si los valores de una columna son diferentes a un valor dado 
        /// </summary>
        /// <param name="table">
        /// Tabla a filtrar
        /// </param>
        /// <param name="column">
        /// Columna para filtrar la tabla
        /// </param>
        /// <param name="value">
        /// Valor por el que se quiere filtrar la tabla
        /// </param>
        /// <returns>
        /// Devuelve la tabla filtrada
        /// </returns>
        public string[,] FilterTableNotEqual(string[,] table, string column, string value)
        {
            int filas = table.GetLength(0);
            int columnas = table.GetLength(1);
            int indiceColumna = -1;

            // Encontrar el índice de la columna
            for (int j = 0; j < columnas; j++)
            {
                if (table[0, j] == column)
                {
                    indiceColumna = j;
                    break;
                }
            }

            if (indiceColumna == -1)
            {
                throw new ArgumentException("Columna no encontrada");
            }

            // Crear una lista para almacenar las filas filtradas
            List<string[]> filasFiltradas = new List<string[]>();

            // Añadir la fila de encabezado
            filasFiltradas.Add(new string[columnas]);
            for (int j = 0; j < columnas; j++)
            {
                filasFiltradas[0][j] = table[0, j];
            }

            // Filtrar las filas según el criterio
            for (int i = 1; i < filas; i++)
            {
                bool agregarFila = false;
                string valor = table[i, indiceColumna];

                if (valor != value)
                {
                    agregarFila = true;
                }

                if (agregarFila)
                {
                    string[] fila = new string[columnas];
                    for (int j = 0; j < columnas; j++)
                    {
                        fila[j] = table[i, j];
                    }
                    filasFiltradas.Add(fila);
                }
            }

            // Convertir la lista de filas filtradas a una matriz bidimensional
            string[,] tablaFiltrada = new string[filasFiltradas.Count, columnas];
            for (int i = 0; i < filasFiltradas.Count; i++)
            {
                for (int j = 0; j < columnas; j++)
                {
                    tablaFiltrada[i, j] = filasFiltradas[i][j];
                }
            }

            return tablaFiltrada;
        }

        /// <summary>
        /// Filtra una tabla dada según si los valores de una columna son mayores o menores que un valor dado 
        /// </summary>
        /// <param name="table">
        /// Tabla a filtrar
        /// </param>
        /// <param name="column">
        /// Columna para filtrar la tabla
        /// </param>
        /// <param name="value">
        /// Valor por el que se quiere filtrar la tabla
        /// </param>
        /// <param name="minor">
        /// Variable opcional para elegir entre "mayor que" o "menor que". Por defecto el valor es false, 
        /// por lo que la función compararía con "mayor que"
        /// </param>
        /// <returns>
        /// Devuelve la tabla filtrada
        /// </returns>
        public string[,] FilterTableByComparison(string[,] table, string column, double value, bool? minor = null)
        {
            int filas = table.GetLength(0);
            int columnas = table.GetLength(1);
            int indiceColumna = -1;

            // Encontrar el índice de la columna
            for (int j = 0; j < columnas; j++)
            {
                if (table[0, j] == column)
                {
                    indiceColumna = j;
                    break;
                }
            }

            if (indiceColumna == -1)
            {
                throw new ArgumentException("Columna no encontrada");
            }

            // Crear una lista para almacenar las filas filtradas
            List<string[]> filasFiltradas = new List<string[]>();

            // Añadir la fila de encabezado
            filasFiltradas.Add(new string[columnas]);
            for (int j = 0; j < columnas; j++)
            {
                filasFiltradas[0][j] = table[0, j];
            }

            // Filtrar las filas según el criterio
            for (int i = 1; i < filas; i++)
            {
                bool agregarFila = false;
                string valor = table[i, indiceColumna];

                if (minor==true)
                {
                    if(double.TryParse(valor, out double numero)&& numero < value)
                    {
                        agregarFila=true;
                    }
                }
                else if(minor == false)
                {
                    if (double.TryParse(valor, out double numero) && numero > value)
                    {
                        agregarFila = true;
                    }
                }

                if (agregarFila)
                {
                    string[] fila = new string[columnas];
                    for (int j = 0; j < columnas; j++)
                    {
                        fila[j] = table[i, j];
                    }
                    filasFiltradas.Add(fila);
                }
            }

            // Convertir la lista de filas filtradas a una matriz bidimensional
            string[,] tablaFiltrada = new string[filasFiltradas.Count, columnas];
            for (int i = 0; i < filasFiltradas.Count; i++)
            {
                for (int j = 0; j < columnas; j++)
                {
                    tablaFiltrada[i, j] = filasFiltradas[i][j];
                }
            }

            return tablaFiltrada;
        }

    }
}