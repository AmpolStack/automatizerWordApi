
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.Json;
using WordApiConverterv;

namespace automatizerWordApi
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            // Add services to the container.
            builder.Services.AddAuthorization();

            builder.Services.AddEndpointsApiExplorer();
            builder.Services.AddSwaggerGen();
            builder.Services.AddCors(options =>
            {
                options.AddPolicy("AllowSpecificOrigin", policy =>
                {
                    policy.WithOrigins("http://localhost:4200") // Permite el origen de Angular
                          .AllowAnyMethod() // Permitir cualquier método (GET, POST, etc.)
                          .AllowAnyHeader(); // Permitir cualquier encabezado
                });
            });
            var app = builder.Build();
            app.UseCors("AllowSpecificOrigin");
            // Configure the HTTP request pipeline.
            if (app.Environment.IsDevelopment())
            {
                app.UseSwagger();
                app.UseSwaggerUI();
            }

            app.UseHttpsRedirection();

            app.UseAuthorization();

            app.MapPost("/giveDocument", ([FromForm] IFormFile file) =>
            {
                string format = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                Dictionary<string, string> keysAndHeaders = new Dictionary<string, string>()
                {
                    { "resolución", "101010"},
                    { "fecha", "20-20-2020" },
                    { "dias" , "10" },
                    { "mes" , "Agosto" }
                };
                string chroma = "@";

                // Validación del archivo
                if (file.ContentType != format)
                {
                    return Results.BadRequest("El archivo no corresponde con la extensión .docx");
                }

                if (file.Length <= 0)
                {
                    return Results.BadRequest("El archivo enviado no tiene contenido alguno");
                }

                // Crear un archivo temporal para almacenar el contenido
                string tempFilePath = Path.GetTempFileName();

                // Copiar el contenido del IFormFile al archivo temporal
                using (var fileStream = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write))
                {
                    file.CopyTo(fileStream); // Sincrónico, ya que no usas async
                }

                // Abrir el archivo temporal para lectura/escritura
                using (var readStream = new FileStream(tempFilePath, FileMode.Open, FileAccess.ReadWrite))
                {
                    wordProcess wordC = new wordProcess();
                    wordC.WordGenerator(readStream, keysAndHeaders, chroma);

                    // Crear un MemoryStream para devolver el archivo generado
                    readStream.Seek(0, SeekOrigin.Begin); // Reiniciar la posición del stream

                    var memoryStream = new MemoryStream();
                    readStream.CopyTo(memoryStream);
                    memoryStream.Seek(0, SeekOrigin.Begin); // Reiniciar para la descarga

                    return Results.File(memoryStream, format, "documentoModificado.docx");
                }

            }).DisableAntiforgery();


            app.MapPost("/findHeaders", ([FromForm] IFormFile file, [FromQuery] string chroma) =>
            {
                if (file.ContentType != "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                {
                    return Results.BadRequest("El archivo no corresponde con la extensión .docx");
                }
                if(file.Length <= 0)
                {
                    return Results.BadRequest("El archivo enviado no tiene contenido alguno");
                }
                using(var stream = file.OpenReadStream())
                {
                    wordProcess wordC = new wordProcess();
                    var wordHeaders = wordC.GetKeysForChroma(stream, chroma);
                    if(wordHeaders.Length > 0)
                    {
                        return Results.Json(new { headers = wordHeaders });
                    }
                    return Results.Ok("El archivo no contiene headers");
                }

            }).DisableAntiforgery();
            app.MapPost("/findHeaderstotal", async(HttpContext httpContext, [FromQuery] string chroma) =>
            {
                var form = await httpContext.Request.ReadFormAsync();
                var file = form.Files.FirstOrDefault();
                if(file == null)
                {
                    return Results.BadRequest("el archivo está vacío");
                }

                if (file.ContentType != "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                {
                    return Results.BadRequest("El archivo no corresponde con la extensión .docx");
                }
                if (file.Length <= 0)
                {
                    return Results.BadRequest("El archivo enviado no tiene contenido alguno");
                }
                using (var stream = file.OpenReadStream())
                {
                    wordProcess wordC = new wordProcess();
                    var wordHeaders = wordC.GetKeysForChroma(stream, chroma);
                    if (wordHeaders.Length > 0)
                    {
                        return Results.Json(new { headers = wordHeaders });
                    }
                    return Results.Ok("El archivo no contiene headers");
                }

            }).DisableAntiforgery();
            app.MapPost("/giveDocumentTotal", async(HttpContext httpContext) =>
            {
                var form = await httpContext.Request.ReadFormAsync();
                var file = form.Files.FirstOrDefault();
                var jsonData = form["data"];
                var data = JsonSerializer.Deserialize<inputHeadersAndValues>(jsonData);
                string format = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                Dictionary<string, string> keysAndHeaders = new Dictionary<string, string>();

                if(data.headers.Length != data.values.Length)
                {
                    return Results.BadRequest("las listas enviadas no tienen el mismo numero de elementos");
                }

                if (data.headers.Length==0 || data.values.Length==0)
                {
                    return Results.BadRequest("las listas no contienen elementos");
                }

                for (int i=0; i<data.headers.Length; i++)
                {
                    keysAndHeaders.Add(data.headers[i], data.values[i]);
                }
                // Validación del archivo
                if (file.ContentType != format)
                {
                    return Results.BadRequest("El archivo no corresponde con la extensión .docx");
                }

                if (file.Length <= 0)
                {
                    return Results.BadRequest("El archivo enviado no tiene contenido alguno");
                }

                // Crear un archivo temporal para almacenar el contenido
                string tempFilePath = Path.GetTempFileName();

                // Copiar el contenido del IFormFile al archivo temporal
                using (var fileStream = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write))
                {
                    file.CopyTo(fileStream); // Sincrónico, ya que no usas async
                }

                var memoryStream = conversionAction(tempFilePath, keysAndHeaders, data.chroma);
                return  Results.File(memoryStream, format, "documentoModificado.docx");
            }).DisableAntiforgery();

            app.MapPost("/giveListDocumentWithZip", async (HttpContext httpContext) =>
            {
                var form = await httpContext.Request.ReadFormAsync();
                var file = form.Files.FirstOrDefault();
                var jsonData = form["data"].ToString(); // Asegúrate de que jsonData esté en formato string

                var data = JsonSerializer.Deserialize<inputHeadersAndValuesList>(jsonData);
                string format = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

                

                if (file?.ContentType != format)
                {
                    return Results.BadRequest("El archivo no corresponde con la extensión .docx");
                }

                List<MemoryStream> fileToConvertZip = new List<MemoryStream>();
                var nameValues = new List<string>();
                foreach (var itemDocument in data.values)
                {
                    nameValues.Add(itemDocument[data.indexforName]);
                    Dictionary<string, string> keysAndHeaders = new Dictionary<string, string>();
                    if (data.headers.Length == 0 || itemDocument.Length == 0)
                    {
                        return Results.BadRequest("Las listas no contienen elementos");
                    }
                    if (data.headers.Length != itemDocument.Length)
                    {
                        return Results.BadRequest("Las listas enviadas no tienen el mismo número de elementos");
                    }
                    for (int i = 0; i < itemDocument.Length; i++)
                    {
                        keysAndHeaders.Add(data.headers[i], itemDocument[i]);
                    }

                    string tempFilePath = Path.GetTempFileName();

                    using (var fileStream = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write))
                    {
                        await file.CopyToAsync(fileStream);
                    }

                    var document = conversionAction(tempFilePath, keysAndHeaders, data.chroma);
                    fileToConvertZip.Add(document);
                }
                
                var zipfile = CreateZipFromMemoryStreams(fileToConvertZip, nameValues.ToArray(),data.customName);
                
                // Crear un ZIP usando ZipOutputStream
                return Results.File(zipfile, "application/zip", "document.zip");



            }).DisableAntiforgery();

            app.Run();
        }
        public static MemoryStream conversionAction(string tempFilePath,  Dictionary<string,string> keysAndHeaders, string chroma)
        {
            using (var readStream = new FileStream(tempFilePath, FileMode.Open, FileAccess.ReadWrite))
            {
                wordProcess wordC = new wordProcess();
                wordC.WordGenerator(readStream, keysAndHeaders, chroma);

                // Crear un MemoryStream para devolver el archivo generado
                readStream.Seek(0, SeekOrigin.Begin); // Reiniciar la posición del stream

                var memoryStream = new MemoryStream();
                readStream.CopyTo(memoryStream);
                memoryStream.Seek(0, SeekOrigin.Begin); // Reiniciar para la descarga

                return memoryStream;
            }

        }
        public static Stream CreateZipFromMemoryStreams(List<MemoryStream> memoryStreams, string[] names, string customName){
        // Creamos un MemoryStream para almacenar el ZIP
        var zipStream = new MemoryStream();

            // Creamos un ZipOutputStream utilizando el MemoryStream
            var zipOutputStream = new ZipOutputStream(zipStream);
        
            zipOutputStream.SetLevel(3); // Establece el nivel de compresión (0-9)
            int counter = 0;
            // Iteramos sobre cada MemoryStream en la lista
            foreach (var memoryStream in memoryStreams)
            {
                // Asegúrate de que el MemoryStream no esté cerrado
                if (memoryStream == null || memoryStream.Length == 0)
                {
                    throw new ArgumentException("Uno de los MemoryStreams está vacío o nulo.");
                }

                // Generamos un nombre único para cada archivo dentro del ZIP
                string entryName = customName.Trim() + names[counter] + ".docx";
                counter++;
                // Creamos una nueva entrada ZIP
                var newEntry = new ZipEntry(entryName)
                {
                    DateTime = DateTime.Now,
                    Size = memoryStream.Length
                };

                // Agregamos la entrada al ZipOutputStream
                zipOutputStream.PutNextEntry(newEntry);

                // Restablecemos la posición del MemoryStream antes de copiarlo al ZIP
                memoryStream.Position = 0;
                memoryStream.CopyTo(zipOutputStream);

                // Cerramos la entrada
                zipOutputStream.CloseEntry();
            }

            // Finalizamos el ZIP
            zipOutputStream.Finish();
        

        // Restablecemos la posición del MemoryStream para que pueda ser leído al devolverlo
        zipStream.Position = 0;

        return zipStream; // Retornamos el MemoryStream que contiene el ZIP
    }

        //public record inputHeadersAndValues(string[] headers, string[] values);
        public class inputHeadersAndValues
        {
            public string[] headers { get; set; }
            public string[] values { get; set; }
            public string chroma {get; set;}
        }
        public class inputHeadersAndValuesList
        {
            public string[] headers { get; set; }
            public string[][] values { get; set; }
            public string chroma { get; set; }
            public int indexforName { get; set; }
            public string customName { get; set; }

        }



    }
}
