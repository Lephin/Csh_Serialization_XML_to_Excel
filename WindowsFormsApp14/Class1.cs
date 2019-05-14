using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;         // Подключаем библиотеку Excel
using System.Xml.Linq;                        // Подключаем библиотеку XML
using System.IO;

namespace WindowsFormsApp14
{

    [Serializable()]
    [XmlRoot(Namespace = "", IsNullable = false, ElementName = "art_list")]
    
	public partial class Art_list
    {

        /// <summary>Список элементов</summary>
        [XmlElement("art_item")]
        //Art_listArt_item - тип свойства Art_item - имя свойства  { get; set; } - объявление методов автосвойства доступного для чтения и записи
		public Art_listArt_item[] Art_item { get; set; } 
       
        /// <summary>Метод создающий новую книгу и сохраняющий в ней данные списка</summary>
        public void SaveNewBook()
        {
            Application excelApp = new Application();
            excelApp.Visible = true;
            Workbook workbook = excelApp.Workbooks.Add(System.Reflection.Missing.Value);
            Worksheet woorksheet = workbook.Sheets[1];

            int currRow = 1;

            woorksheet.Cells[currRow, 1].Value2 = "ID";
            woorksheet.Cells[currRow, 2].Value2 = "Обозначение";
            woorksheet.Cells[currRow, 3].Value2 = "Наименование";
            woorksheet.Cells[currRow, 4].Value2 = "Раздел";
            woorksheet.Cells[currRow, 5].Value2 = "Код исполнения";
            woorksheet.Cells[currRow, 6].Value2 = "Покупной";
            woorksheet.Cells[currRow, 7].Value2 = "Масса";
            woorksheet.Cells[currRow, 8].Value2 = "Дата создания";
            woorksheet.Cells[currRow, 9].Value2 = "Серийный номер";

            currRow++;

            //ToXL это метод он находится в Class2
            foreach (Art_listArt_item item in Art_item) 
                item.ToXLS(woorksheet.Cells[currRow++, 1]);

            woorksheet.Cells[currRow, 6].Value2 = "ИТОГО:";
            woorksheet.Cells[currRow, 7].FormulaR1C1 = "=SUM(R1C:R[-1]C)";

        }

        public string Print()
            => string.Join("\r\n", Art_item.Select(item => item.Print()));
    }
}
