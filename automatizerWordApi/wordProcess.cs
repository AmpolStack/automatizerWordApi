using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.AspNetCore.Http.HttpResults;
using System.Runtime.InteropServices.Marshalling;
using System.Text.RegularExpressions;
namespace WordApiConverterv
{
    public class wordProcess
    {
        public void WordGenerator(Stream templatePath, Dictionary<string, string> data, string chroma)
        {
            
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(templatePath, true))
                {
                    // Obtener el cuerpo principal del documento
                    var body = wordDoc.MainDocumentPart.Document.Body;
                    var text = body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>();
                    var ListText = text.ToList();
                int counter = 0;
                    //Recorrer cada entrada en el diccionario de datos
                    foreach (var entry in data)
                    {
                    
                        for (int i = 0; i < ListText.Count; i++)
                        {
                        Console.WriteLine(ListText[i].Text);

                        //Console.WriteLine("item : " + ListText[i].Text);
                        if (Regex.IsMatch(ListText[i].Text, @$"\{chroma}{entry.Key}\{chroma}"))
                            {


                            ListText[i].Text = Regex.Replace(ListText[i].Text, $@"\{chroma}{entry.Key}\{chroma}", entry.Value);
                            counter++;
                            //Console.WriteLine("match = " + data.Keys.Count + "\nreplaces total = " + counter);
                            }
                            if (i >= ListText.Count - 2)
                            {
                            Console.WriteLine("Aquí me detuve");
                            continue;
                            }
                            if (ListText[i + 1].Text.Contains(entry.Key) && ListText[i + 2].Text.Contains(chroma) && ListText[i].Text.Contains(chroma))
                            {
                                ListText[i].Text = "";
                                ListText[i + 1].Text = entry.Value;
                                ListText[i + 2].Text = "";
                            }

                        }
                    }
                    text = ListText;
                    wordDoc.MainDocumentPart.Document.Save();
                    
                }
        }
        public string[] GetKeysForChroma(Stream templatePath, string chroma)
        {
            // Crear una lista para almacenar las palabras encontradas
            List<string> findWords = new List<string>();
            

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(templatePath, false)) // Modo lectura del docx asignación
            {
                // Obtener el texto del .docx
                string docText = null;

                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                // Eliminar cualquier etiqueta de estilo HTML o XML (si es necesario)
                docText = Regex.Replace(docText, @"<\s*\/?\s*([a-zA-Z_][\w.-]*)\s*[^>]*>", "");

                // Buscar las coincidencias con el patrón usando chroma1 y chroma2
                MatchCollection matches = Regex.Matches(docText, @$"\{chroma}\w+\{chroma}");

                // Iterar sobre cada coincidencia
                foreach (Match item in matches)
                {
                    string matchValue = item.Groups[0].Value;

                    // Solo agregar si el elemento no está ya en la lista
                    if (!findWords.Contains(matchValue))
                    {
                        findWords.Add(matchValue);
                    }
                }

                // Retornar los resultados como un array de strings
                return findWords.ToArray();
            }
        }

    }
}
