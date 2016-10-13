using System;
using System.IO;


namespace SammamishMeterImport
{
  class errorlogging
  {

    public void logMessage(string LogFilePathAndName, string message)
    {

      if (!File.Exists(LogFilePathAndName))
      {
        // Create a file to write to.
        StreamWriter swNew = File.CreateText(LogFilePathAndName);
        swNew.WriteLine(DateTime.Now.ToString()+": " + message);
        swNew.Close();
      }
      else
      {
        StreamWriter swAppend = File.AppendText(LogFilePathAndName);
        swAppend.WriteLine(message);
        swAppend.Close();
      }
    }
   }
}
