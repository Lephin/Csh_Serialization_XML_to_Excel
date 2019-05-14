using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Xml.Serialization;

namespace WindowsFormsApp14
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e) //Событие Клик для кнпоки
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return; // Ищем файл
            label1.Text = openFileDialog1.SafeFileName; // Выводим название выбранного файла в компонент

            string fileName = openFileDialog1.FileName;  // Путь к файлу передали в переменную

            Art_list data;

            try
            {
                using (FileStream fs = new FileStream(fileName, FileMode.Open))// Файловой поток
                {
                    XmlSerializer x = new XmlSerializer(typeof(Art_list)); //Создание сериализатор по заданному классу 
                    
					data = (Art_list)x.Deserialize(fs); //Получение данных и создание по ним объекта
                    fs.Close();
                }
            }
            catch (Exception)
            {
                throw new ArgumentException("Ошибка десериализации класса art_list из файла " + fileName);
            }

            //Метод сохранения находится в Class1
            data.SaveNewBook();
           
        }
    }
}
