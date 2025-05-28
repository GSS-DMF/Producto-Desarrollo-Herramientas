using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using Microsoft.Office.Interop.Word;



namespace RepositorioFuncionesGitHub
{
    public class MSWord
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



        /// <summary>
        /// Crea un documento nuevo a partir de un documento original que se toma como plantilla.
        /// El documento original está en una ruta absoluta y tiene predefinido un formato de estilos.
        /// El documento creado se guarda en una ruta especificada con un nombre espeficicado.
        /// </summary>
        /// <param name="fileName">
        /// Nombre con el que se quiere guardar el documento Word
        /// </param>
        /// <param name="wordPath">
        /// Ruta de la carpeta en la que se quiere guardar el documento de Word
        /// </param>
        /// <param name="templatePath">
        /// Ruta del word plantilla (con nombre de plantilla incluido, y extensión .docx
        /// </param>
        /// <returns>
        /// Ruta del archivo guardado, con nombre de archivo y extensión .docx
        /// </returns>
        public string CreateDocument(string fileName, string wordPath, string templatePath)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            object missing = System.Reflection.Missing.Value;
            object readOnly = false;
            object isVisible = true;
            object filePath = templatePath;
            object fileName2 = "";

            try
            {
                Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(ref filePath, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);

                // Guardar y cerrar el documento

                fileName2 = System.IO.Path.Combine(wordPath, fileName);
                object fileFormat = WdSaveFormat.wdFormatDocumentDefault;

                try
                {
                    doc.SaveAs2(ref fileName2, ref fileFormat, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                    //MessageBox.Show($"Documento guardado como: {fileName2}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al guardar el documento: {ex.Message}");
                }

                doc.Close();
                wordApp.Quit();

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al abrir el documento: {ex.Message}");
            }

            return fileName2.ToString() + ".docx";
        }

        /// <summary>
        /// Abrir un word existente, dada una ruta
        /// </summary>
        /// <param name="wordPath">
        /// Ruta de documento, incluyendo nombre y extensión ".docx"
        /// </param>
        /// <returns>
        /// Objeto word abierto
        /// </returns>
        public Microsoft.Office.Interop.Word.Document OpenWord(string wordPath)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            object missing = System.Reflection.Missing.Value;
            object readOnly = false;
            object isVisible = true;
            object fileName = wordPath;
            Microsoft.Office.Interop.Word.Document doc = null;
            try
            {
                doc = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al abrir el documento: {ex.Message}");
            }
            return doc;
        }

        /// <summary>
        /// Cierra el documento de word abierto y la aplicación
        /// </summary>
        /// <param name="doc">
        /// Objeto Word abierto
        /// </param>
        public void CloseWord(Microsoft.Office.Interop.Word.Document doc)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            try
            {
                doc.Save();
                doc.Close();
                Marshal.ReleaseComObject(doc);
                Marshal.ReleaseComObject(wordApp);
                wordApp.Quit();
                doc = null;
                wordApp = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cerrar el documento: {ex.Message}");
            }
        }

        /// <summary>
        /// Añade un texto al final de un documento word seleccionado. Tiene opción de añadirlo como
        /// texto normal o como título
        /// </summary>
        /// <param name="text">
        /// Texto que se quiere añadir al documento
        /// </param>
        /// <param name="doc">
        /// Objeto word abierto
        /// </param>
        /// <param name="titleStyle">
        /// Booleano para elegir si el texto es normal(false) o título(true). Por defecto es texto normal. 
        /// </param>
        public void AddText(string text, Microsoft.Office.Interop.Word.Document doc, bool? titleStyle = false)
        {
            try
            {
                // Copiar el texto almacenado en la variable 'texto' al documento de Word
                Microsoft.Office.Interop.Word.Range range = doc.Content;
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.Font.Name = "Neo Tech Std";
                range.Text += text;
                if (titleStyle == true)
                {
                    range.set_Style(WdBuiltinStyle.wdStyleHeading1);
                }
                else
                {
                    range.set_Style(WdBuiltinStyle.wdStyleNormal);
                }
                range.Text += "\n";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al abrir el documento: {ex.Message}");
            }
        }

        /// <summary>
        /// Añade un salto de página al final de un documento word seleccionado. 
        /// </summary>
        /// <param name="doc">
        /// Objeto word abierto
        /// </param>
        public void AddPageBreak(Microsoft.Office.Interop.Word.Document doc)
        {
            try
            {
                // Copiar el texto almacenado en la variable 'texto' al documento de Word
                Microsoft.Office.Interop.Word.Range range = doc.Content;
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.InsertBreak(WdBreakType.wdPageBreak);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al abrir el documento: {ex.Message}");
            }
        }

        /// <summary>
        /// Añade una tabla a un documento word seleccionado.
        /// </summary>
        /// <param name="table">
        /// Tabla a añadir al documento word
        /// </param>
        /// <param name="doc">
        /// Objeto word abierto
        /// </param>
        public void AddTable(string[,] table, Microsoft.Office.Interop.Word.Document doc)
        {
            try
            {
                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                object missing = System.Reflection.Missing.Value;

                //Desactivar la actualización de la pantalla
                wordApp.ScreenUpdating = false;

                // Añadir una tabla al final del documento
                Microsoft.Office.Interop.Word.Range range = doc.Content;
                range.Collapse(WdCollapseDirection.wdCollapseEnd);

                Table wordtable = doc.Tables.Add(range, table.GetLength(0), table.GetLength(1), ref missing, ref missing);
                //table.Borders.Enable = 1; //Mejor sin bordes

                //Aplicar formato general a la tabla
                wordtable.Range.Font.Name = "Neo Tech Std";
                wordtable.Range.Font.Size = 8;
                wordtable.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                wordtable.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                for (int j = 0; j < table.GetLength(1); j++)
                {
                    wordtable.Cell(1, j + 1).Range.Font.Bold = 1;
                    wordtable.Cell(1, j + 1).Shading.BackgroundPatternColor = (WdColor)(0xF3E2D9);
                }

                // Rellenar la tabla con datos
                for (int i = 0; i < table.GetLength(0); i++)
                {
                    for (int j = 0; j < table.GetLength(1); j++)
                    {
                        wordtable.Cell(i + 1, j + 1).Range.Text = table[i, j];
                    }
                }

                //Reactivar la actualización de pantalla
                wordApp.ScreenUpdating = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al abrir el documento: {ex.Message}");
            }
        }

        /// <summary>
        /// Pegar contenido en un documento de word
        /// </summary>
        /// <param name="doc">
        /// Objeto Word abierto
        /// </param>
        public void Paste(Microsoft.Office.Interop.Word.Document doc)
        {
            try
            {
                Microsoft.Office.Interop.Word.Range range = doc.Content;
                range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                range.Paste();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"No se encuentra el documento: {ex.Message}");
            }
        }

        /// <summary>
        /// Ajusta el ancho de la última tabla añadida al ancho de la página
        /// </summary>
        public void AutoFitTableWidth(Microsoft.Office.Interop.Word.Document doc)
        {
            try
            {
                //Ajustar el ancho de la tabla al ancho de la página
                if (doc.Tables.Count > 0)
                {
                    Microsoft.Office.Interop.Word.Table table = doc.Tables[doc.Tables.Count];
                    table.AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"No se encuentra el documento: {ex.Message}");
            }
        }

        /// <summary>
        /// Da formato de tipo encabezado a las primeras filas de la última tabla del documento seleccionado
        /// </summary>
        /// <param name="doc">
        /// Objeto Word abierto
        /// </param>
        /// <param name="numberRows">
        /// Número de filas a las que se les va a aplicar el formato tipo encabezado.
        /// </param>
        public void FormatHeaderRow(Microsoft.Office.Interop.Word.Document doc, int numberRows)
        {
            try
            {
                if (doc.Tables.Count > 0)
                {
                    Microsoft.Office.Interop.Word.Table table = doc.Tables[doc.Tables.Count];

                    for (int i = 0; i < numberRows; i++)
                    {
                        Microsoft.Office.Interop.Word.Row row = table.Rows[i];
                        row.HeadingFormat = -1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"No se encuentra el documento: {ex.Message}");
            }
        }



        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------

        // Escribir todas las subclases aquí. Añadir un docstring para tener 
        // información acerca de su funcionamiento y parámetros de entrada
        // y salida. Recordar añadirlo al excel de registro de métodos. Poner 
        // todos los métodos públicos para evitar errores de acceso.

        //---------------------------------------------------------------------------------
        //---------------------------------------------------------------------------------




    }
}