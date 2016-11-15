using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.ComponentModel;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace XlsLibrary
{
    //com component
    [ComVisible(true)]
    public interface IConverterInterface
    {
        void Convert();
    }

    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class Converter : IConverterInterface
    {
        public Converter()
        {
        }

        //[ComVisible(true)]
        public void Convert()
        {
            Excel.Application xlApp = new Excel.Application();

            Excel.Workbook xlWorkBook;

            //~~> Start Excel and open the workbook.
            xlWorkBook = xlApp.Workbooks.Open("C:\\temp\\book1.xls");

            //~~> Run the macros by supplying the necessary arguments
            xlApp.Run("Macro1", "1", "2");

            //~~> Clean-up: Close the workbook
            xlWorkBook.Close(false);

            //~~> Quit the Excel Application
            xlApp.Quit();

            //~~> Clean Up
            releaseObject(xlApp);
            releaseObject(xlWorkBook);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
