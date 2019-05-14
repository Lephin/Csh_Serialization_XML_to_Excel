using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel; // Подключаем библиотеку Excel
using System.Xml.Linq;                // Подключаем библиотеку XML
using System.IO;
using System.Globalization;

namespace WindowsFormsApp14

{
    /// <summary>Класс для элементов списка</summary>
    [Serializable()]

    public partial class Art_listArt_item
    {
        #region Секция сереализуемых свойств
        [XmlElement("attr_item")] // Переменная будет иметь значения элемента 
        public Art_listArt_itemAttr_item[] Attr_item { get; set; }
        
		[XmlAttribute("art_id")] // Переменная будет иметь значения атрибута
        public byte Art_id { get; set; }
        
		[XmlAttribute("designation")] // Переменная будет иметь значения атрибута
        public string Designation { get; set; }
        
		[XmlAttribute("name")] // Переменная будет иметь значения атрибута
        public string Name { get; set; }
        
		[XmlAttribute("section_id")]
        public byte Section_id { get; set; }
        #endregion

        #region Секция свойств сериализуемых в список Attr_item
        public object GetValue(string NameValue)
        {
            Type typeValue = Types[NameValue];
            object ret;
            
			try
            {
                if (typeValue == typeof(DateTime))
                    ret = Convert.ChangeType
                    (
                        Attr_item.FirstOrDefault(val => val.Attr_name == NameValue).Value,
                        typeValue,
                        new CultureInfo("Ru-ru")
                    );
                else
                    ret = Convert.ChangeType
                    (
                        Attr_item.FirstOrDefault(val => val.Attr_name == NameValue).Value,
                        typeValue,
                        CultureInfo.InvariantCulture
                    );
            }
            catch (Exception)
            {
                ret = null;
            }
            return ret;
        }

        public static Dictionary<string, Type> Types = new Dictionary<string, Type>()
        {
            {"Код исполнения",typeof(int) },
            {"Покупной",typeof(string) },
            {"Масса",typeof(decimal) },
            {"Дата создания",typeof(DateTime) },
            {"Серийный номер",typeof(string) },
        };
        
		/// <summary>Код исполнения</summary>
        public int? ExecutionCode => (int?)GetValue("Код исполнения");
        /// <summary>Покупной</summary>
        public bool? Purchased
        {
            get
            {
                string val = (string)GetValue("Покупной");
                return string.IsNullOrWhiteSpace(val) ? null : val.StartsWith("+") ? true : val.StartsWith("-") ? (bool?)false : null;
            }
        }
        
		/// <summary>Масса</summary>
        public decimal? Weight => (decimal?)GetValue("Масса");
        /// <summary>Дата создани</summary>
        public DateTime? DateOfCreation => (DateTime?)GetValue("Дата создания");
        /// <summary>Серийный номер</summary>
        public string SerialNumber => (string)GetValue("Серийный номер");
        #endregion

        /// <summary>Метод сохранения данных класса на лист Excel в одну строку с указанной ячейки</summary>
        /// <param name="cell">Указанная ячейка</param>
		
		
        public void ToXLS(Range cell)
        {

            Range _cell = cell[1, 1];
            Worksheet _sheet = _cell.Worksheet;
            
			int row = _cell.Row;
            int col = _cell.Column;

            _sheet.Cells[row, col].Value2 = Art_id;
            _sheet.Cells[row, col + 1].Value2 = Designation;
            _sheet.Cells[row, col + 2].Value2 = Name;
            _sheet.Cells[row, col + 3].Value2 = Section_id;
            
			if (ExecutionCode != null)
                _sheet.Cells[row, col + 4].Value2 = ExecutionCode;
            
			if (Purchased != null)
                _sheet.Cells[row, col + 5].Value2 = Purchased.Value ? "+" : "-";
            
			if (Weight != null)
                _sheet.Cells[row, col + 6].Value2 = Weight;
            
			if (DateOfCreation != null)
                _sheet.Cells[row, col + 7].Value2 = DateOfCreation?.ToString();
            
			if (SerialNumber != null)
                _sheet.Cells[row, col + 8].Value2 = SerialNumber;

            return;
        }


        public string Print()
            => $"Art_listArt_item: Art_id={Art_id}, Designation={Designation}, Name={Name}, Section_id={Section_id}\r\n\t" +
            string.Join("\r\n\t", Attr_item.Select(item => item.Print()));
    }
}
