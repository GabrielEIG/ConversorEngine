using Microsoft.VisualBasic;
using System.Reflection.Metadata;
using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Globalization;

namespace ConversorEngine
{
	internal class Program
	{
		static void Main(string[] args)
		{
			// Ruta de los archivos
			string inputRtfPath = @"ruta de tu archivo .rtf";
			string outputPdfPath = @"ruta donde se descargara tu archivo pdf";
			string imagePath = @"imagen de tu prederencia .jpg"; // Imagen a agregar al final

			try
			{
				// 1. Cargar el archivo RTF
				var doc = new Aspose.Words.Document(inputRtfPath);

				// 2. Crear una nueva sección al inicio con texto
				var builder = new DocumentBuilder(doc);
				builder.MoveToDocumentEnd();

				// Configurar estilo de la primera página
				builder.InsertBreak(BreakType.SectionBreakNewPage);
				builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
				builder.Font.Size = 16;
				builder.Font.Bold = true;
				builder.Writeln("El equipo de desarrollo, son los mejores");

				// 3. Crear una nueva sección al final con una imagen
				builder.MoveToDocumentEnd();
				builder.InsertBreak(BreakType.SectionBreakNewPage);

				builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
				builder.Writeln("Página con una imagen:");

				// Insertar imagen desde archivo
				builder.InsertImage(imagePath, 300, 200); // Ajustar tamaño de la imagen si es necesario

				// 4. Guardar el documento como PDF
				doc.Save(outputPdfPath, SaveFormat.Pdf);

				Console.WriteLine($"El archivo PDF se creó correctamente en: {outputPdfPath}");
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Ocurrió un error: {ex.Message}");
			}
		}
	}
}
