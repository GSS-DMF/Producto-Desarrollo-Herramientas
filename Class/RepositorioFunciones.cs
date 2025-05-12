using System;
using RepositorioFuncionesGitHub;


namespace RepositorioFuncionesGitHub
{
    public class RepositorioFunciones
    {
        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------

        // Clase auxiliar que contiene todas las instancias de las clases de SAP, Excel, 
        // Word... Cada vez que creemos una nueva clase, debemos instanciarla aqui.

        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------
        
        public MSWindows MSWindows = new MSWindows();
        public MSExcel MSExcel = new MSExcel();
        public MSWord MSWord = new MSWord();
        public SAP SAP = new SAP();
        public Math Math = new Math();
        public Tables Tables = new Tables();

    }
}