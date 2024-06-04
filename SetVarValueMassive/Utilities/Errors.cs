using ExcelDataReader.Log;
using System;
using System.Collections.Generic;

namespace SetVarValueMassive.Utilities
{
    public static class Errors
    {
        public static List<string> mensajesExcel = new List<string>();
        public static List<string> mensajesSeteo = new List<string>();
        public static List<string> mensajesVault = new List<string>();        
        
        //Metodo para mostrar por consola un mensaje 
        public static void ShowMessage(string message, string encabezado)
        {
            if(encabezado == "EXCELCONTROLLER")
            {
                mensajesExcel.Add(message);
            }
            if (encabezado == "SETVARVALUES")
            {
                mensajesSeteo.Add(message);
            }
            if (encabezado == "VAULTCONTROLLER")
            {
                mensajesVault.Add(message);
            }
            if ((message == "Seteo de variables de carpeta realizado correctamente") ||
                message == "Seteo de variables de archivos realizado correctamente" ||
                message.Contains("No se encontro una columna referida a los ID de los archivos"))
            {
                ShowConsoleMessage(mensajesVault, "VAULTCONTROLLER");
                ShowConsoleMessage(mensajesExcel, "EXCELCONTROLLER");
                ShowConsoleMessage(mensajesSeteo, "SETEOCONTROLLER");
            }
        }

        private static void ShowConsoleMessage(List<string> mensajesAMostrar, string encabezado)
        {
            Console.WriteLine($"{Environment.NewLine}<<<===============((  OPERACIONES CONTROLADOR: {encabezado}  ))===============>>> {Environment.NewLine}");
            foreach (string mensaje in mensajesAMostrar)
            {
                ConsoleColor color = mensaje.Contains("Error") ? ConsoleColor.Red : ConsoleColor.Green;
                Console.ForegroundColor = color;
                Console.WriteLine(mensaje);
            }
            Console.ResetColor();
            Console.WriteLine($"{Environment.NewLine}<<<===================================\\\\\\\\////===================================>>> {Environment.NewLine}");
        }
    }
}
