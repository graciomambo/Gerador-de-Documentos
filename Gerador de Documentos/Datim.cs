using LinqToExcel;
using LinqToExcel.Attributes;
using LinqToExcel.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Gerador_de_Documentos
    {
    class Datim
        {
      
        public String Id { get; set; }
        
        public String Desagregado { get; set; }

       
        public String SuportType { get; set; }
       
        public String EntryField { get; set; }
       
        public String Value { get; set; }
       
        public String DataSet { get; set; }
      
        public String Distrito { get; set; }
        
        public String US { get; set; }


        public String SourceTable { get; set; }

      
        public String Indicators { get; set; }
        
      
       
        public string CAMPO_INDICADOR_DEFEITO = "Source_Table";
        public string CAMPO_US_DEFEITO;
        public string CAMPO_VALOR_DEFEITO = "Value";
        public string FILTRO_INDICADOR_DEFEITO = "Todos";
        public string FILTRO_US_DEFEITO = "Todas";
        public string FILTRO_VALOR_DEFEITO = "Excluir Zero";
        public string SHEET_DEFEITO = "sheet1";
 
        

        
        }
    }

