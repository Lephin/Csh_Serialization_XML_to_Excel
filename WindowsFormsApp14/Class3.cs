using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;
using System.Xml;
using System.Xml.Serialization;

namespace WindowsFormsApp14
{
    /// <summary>Вспомогательный класс для XML сериализации</summary>
    [Serializable()]
    [DesignerCategory("code")]
    [XmlType(AnonymousType = true)]

    public partial class Art_listArt_itemAttr_item
    {

        /// <remarks/>
        [XmlAttribute("attr_name")]
        public string Attr_name { get; set; }

        /// <remarks/>
        [XmlText()]
        public string Value { get; set; }

        public string Print()
            => $" Art_listArt_itemAttr_item: Attr_name={Attr_name}, Value={Value}";
    }
}
