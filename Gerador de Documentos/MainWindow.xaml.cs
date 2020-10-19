

using LinqToExcel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Gerador_de_Documentos;
using LinqToExcel.Query;
using System.Threading;
using LinqToExcel.Extensions;
using System.Runtime.Remoting.Metadata.W3cXsd2001;

namespace Gerador_de_Documentos
    {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 

    public partial class MainWindow : Window
        {
        const String SELECCAO_VAZIA = "Nenhum Ficheiro Selecionado";
        
        String filePath;
        Datim datim= new Datim();
        private ExcelQueryFactory excel;
        private List<string> listaColunas;
        private List<string> distinctUS;
        private List<string> distinctSourceTable;

        public MainWindow()
            {
            InitializeComponent();
            ficheirosSelecionados.Items.Add(SELECCAO_VAZIA);
            comboIndicadorListaCampo.IsEnabled=false;
            comboIndicadorListaFiltros.IsEnabled=false;
            comboUSListaCampo.IsEnabled=false;
            comboUSListaFiltros.IsEnabled=false;
            comboValueListaCampo.IsEnabled=false;
            comboValueListaFiltros.IsEnabled=false;

            }
        private void clicarSelecionarFicheiro(object sender , RoutedEventArgs e)
            {
        
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect=false;
            openFileDialog.Filter="Excel Files|*.xls;*.xlsx;*.xlsm;*.csv";
            openFileDialog.InitialDirectory=Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
      
            if(openFileDialog.ShowDialog()==true)
                {
                if(ficheirosSelecionados.Items.Count>0)
                    ficheirosSelecionados.Items.Remove(ficheirosSelecionados.Items[0]);
                foreach(string file in openFileDialog.FileNames)
                    ficheirosSelecionados.Items.Add(file);
                filePath=openFileDialog.FileNames[0];
                
                excel=new ExcelQueryFactory(filePath);
               
                excel.ReadOnly = true;
                
                excel.StrictMapping=StrictMappingType.ClassStrict;

                listaColunas = excel.GetColumnNames(datim.SHEET_DEFEITO).ToList();
                //listaColunas = new List<String>();
                //listaColunas.Insert(0, "ID_DATIM");
                //listaColunas.Insert(1, "IndicatorMapping#Desag");
                //listaColunas.Insert(2, "Support_Type");
                //listaColunas.Insert(3, "IndicatorMapping#nteryFildID");
                //listaColunas.Insert(4, "Value");
                //listaColunas.Insert(5, "DataSet");
                //listaColunas.Insert(6, "Distrito_DATIM");
                //listaColunas.Insert(7, "Nome_US_DATIM");
                //listaColunas.Insert(8, "Source#Table");
                //listaColunas.Insert(9, "Indicators");

                excel.AddMapping<Datim>(x => x.Id, listaColunas[0]);
                excel.AddMapping<Datim>(x => x.Desagregado, listaColunas[1]);
                excel.AddMapping<Datim>(x => x.SuportType, listaColunas[2]);
                excel.AddMapping<Datim>(x => x.EntryField, listaColunas[3]);
                excel.AddMapping<Datim>(x => x.Value, listaColunas[4]);
                excel.AddMapping<Datim>(x => x.DataSet, listaColunas[5]);
                excel.AddMapping<Datim>(x => x.Distrito, listaColunas[6]);
                excel.AddMapping<Datim>(x => x.US, listaColunas[7]);
                excel.AddMapping<Datim>(x => x.SourceTable, listaColunas[8]);
                excel.AddMapping<Datim>(x => x.Indicators, listaColunas[9]);


        
             
                datim.CAMPO_US_DEFEITO= listaColunas[7];
                datim.CAMPO_INDICADOR_DEFEITO = listaColunas[8];

                excel.UsePersistentConnection=false;
                
                carregarCampos(sender,e);

                }

            }
        private  void carregarCampos(object sender , RoutedEventArgs e)
            {
            
             Task.Run(() =>
             {
                 Dispatcher.Invoke(() =>
                 {
                 comboIndicadorListaCampo.Items.Clear();
                 comboUSListaCampo.Items.Clear();
                 comboValueListaCampo.Items.Clear();
                     foreach(var lc in listaColunas)
                         {
                         comboIndicadorListaCampo.Items.Add(lc);
                         
                         comboUSListaCampo.Items.Add(lc);
                         comboValueListaCampo.Items.Add(lc);
                         comboIndicadorListaCampo.IsEnabled=true;
                         comboUSListaCampo.IsEnabled=true;
                         comboValueListaCampo.IsEnabled=true;
                         }
                     comboIndicadorListaCampo.SelectedItem=datim.CAMPO_INDICADOR_DEFEITO;
                     comboUSListaCampo.SelectedItem=datim.CAMPO_US_DEFEITO;
                     comboValueListaCampo.SelectedItem=datim.CAMPO_VALOR_DEFEITO;
                 });

             });

             Task.Run(() =>
            {
                distinctSourceTable=(from d in excel.Worksheet<Datim>()
                            select d.SourceTable).Distinct().ToList();
                preencherComboBoxFiltro(comboIndicadorListaFiltros , distinctSourceTable, datim.FILTRO_INDICADOR_DEFEITO);
            });
            Task.Run(() =>
           {
           distinctUS=(from d in excel.Worksheet<Datim>()
                       select d.US).Distinct().ToList();
                preencherComboBoxFiltro(comboUSListaFiltros , distinctUS, datim.FILTRO_US_DEFEITO);
            });
             Task.Run(() =>
            {
                preencherComboBoxFiltroValor(comboValueListaFiltros);
            });
               
            comboIndicadorListaFiltros.SelectedItem=datim.FILTRO_INDICADOR_DEFEITO;
            comboUSListaFiltros.SelectedItem=datim.FILTRO_US_DEFEITO;
            comboValueListaFiltros.SelectedItem=datim.FILTRO_VALOR_DEFEITO;
             
          
            }
        private void preencherComboBoxFiltro(ComboBox filtro , List<string> distintos,String criterioPadrao)
            {

           
            Dispatcher.Invoke(() => {
            
            filtro.Items.Clear();
            filtro.Items.Add(criterioPadrao);
            filtro.SelectedItem=criterioPadrao;
                
           
            foreach(String i in distintos)
                {
                 filtro.Items.Add(i); 
                
                }
                filtro.IsEnabled=true;
            });

            }
        private void preencherComboBoxFiltroValor(ComboBox filtro)
            {
            Dispatcher.Invoke(() => {
            filtro.Items.Clear();
            filtro.Items.Add("Excluir Zero");
            filtro.Items.Add("Incluir Zero");
            filtro.SelectedItem=filtro.Items[0];
            filtro.IsEnabled=true;
            });

            }
        private void clicarBotaoGerar(object sender , RoutedEventArgs e)
            {if(validForm())
                    { MessageBox.Show("Preencha todos campos" , "Gerador" , MessageBoxButton.OK , MessageBoxImage.Information); }
                else {

                    
                        gerarFicheiro();
                
                }
            }

        private bool validForm()
            {
            if(   comboIndicadorListaCampo.SelectedItem==null 
                &comboIndicadorListaFiltros.SelectedItem==null 
                &comboUSListaCampo.SelectedItem==null
                &comboUSListaFiltros.SelectedItem==null
                &comboValueListaCampo.SelectedItem==null
                &comboValueListaFiltros.SelectedItem==null &
                 comboIndicadorListaCampo.Items.IsEmpty
                &comboIndicadorListaFiltros.Items.IsEmpty
                &comboUSListaCampo.Items.IsEmpty
                &comboUSListaFiltros.Items.IsEmpty
                &comboValueListaCampo.Items.IsEmpty
                &comboValueListaFiltros.Items.IsEmpty

                )
                return true;
            else return false;
            }

        private void gerarFicheiro()
            {
            Dispatcher.Invoke(() =>
            {
                PBar.Value = 0;
                PBar.Visibility = Visibility.Visible;
                PBarText.Visibility = Visibility.Visible;
                botaoGerar.IsEnabled = false;
                botaoSelecionar.IsEnabled = false;
                comboIndicadorListaCampo.IsEnabled = false;
                comboIndicadorListaFiltros.IsEnabled = false;
                comboUSListaCampo.IsEnabled = false;
                comboUSListaFiltros.IsEnabled = false;
                comboValueListaCampo.IsEnabled = false;
                comboValueListaFiltros.IsEnabled = false;

            });
           
            List<string> fi = new List<string>();
            List<string> fus = new List<string>();
            List<Datim> listaDatimActual = new List<Datim>();
            if(comboIndicadorListaFiltros.SelectedItem.ToString()==datim.FILTRO_INDICADOR_DEFEITO)
                fi=distinctSourceTable;
            else
                fi.Add(comboIndicadorListaFiltros.SelectedItem.ToString());
            
            if(comboUSListaFiltros.SelectedItem.ToString()==datim.FILTRO_US_DEFEITO)
                fus=distinctUS;
            else
                fus.Add(comboUSListaFiltros.SelectedItem.ToString());




            //(from d in excel.Worksheet<Datim>()
            //select d.SourceTable).Distinct().ToList()
            //where d.SourceTable==comboIndicadorListaCampo.SelectedItem.ToString()&&d.US==comboUSListaCampo.SelectedItem.ToString()


            //.Where(x=>fus.Contains(x.US)).Where(x => fi.Contains(x.SourceTable)).ToList();
            object misValue = System.Reflection.Missing.Value;
            Excel.Application xlApp= new Excel.Application();
            Excel.Workbook xlWorkBook= xlApp.Workbooks.Add(misValue);
            
            Excel.Worksheet xlWorkSheet= (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
           
            int linha = 0;
            int progress = 1;
            string set = comboValueListaFiltros.SelectedItem.ToString();
            Task.Run(() =>
            {
            foreach(string i in fi)
                {
                foreach(string u in fus)
                    {
                        
                            listaDatimActual=(from d in excel.Worksheet<Datim>()
                                              where d.US==u&&d.SourceTable==i
                                              select d).ToList();
                        linha = 1;

                        foreach (var rid in listaDatimActual)
                            {
                           

                            if (set.Equals(datim.FILTRO_VALOR_DEFEITO) & rid.Value != null)
                                {if(linha==1)
                                  {
                                        xlApp = new Excel.Application();
                                        xlApp.Visible=false;
                                        misValue = System.Reflection.Missing.Value;
                                        xlWorkBook=xlApp.Workbooks.Add(misValue);
                                        xlWorkSheet=(Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                                        xlWorkSheet.Cells[1][linha]= "ID_DATIM";
                                        xlWorkSheet.Cells[2][linha]= "IndicatorMapping.Desag";
                                        xlWorkSheet.Cells[3][linha]= "Support_Type";
                                        xlWorkSheet.Cells[4][linha]= "IndicatorMapping.EnteryFildID";
                                        xlWorkSheet.Cells[5][linha]= "Value";
                                        xlWorkSheet.Cells[6][linha]= "DataSet";
                                        xlWorkSheet.Cells[7][linha]= "Distrito_DATIM";
                                        xlWorkSheet.Cells[8][linha]= "Nome_US_DATIM";
                                        xlWorkSheet.Cells[9][linha]= "Source.Table";
                                        xlWorkSheet.Cells[10][linha]= "Indicators";
                                    linha++;
                                        }
                                xlWorkSheet.Cells[1][linha] = rid.Id;
                                xlWorkSheet.Cells[2][linha] = rid.Desagregado;
                                xlWorkSheet.Cells[3][linha] = rid.SuportType;
                                xlWorkSheet.Cells[4][linha] = rid.EntryField;
                                xlWorkSheet.Cells[5][linha] = rid.Value;
                                xlWorkSheet.Cells[6][linha] = rid.DataSet;
                                xlWorkSheet.Cells[7][linha] = rid.Distrito;
                                xlWorkSheet.Cells[8][linha] = rid.US;
                                xlWorkSheet.Cells[9][linha] = rid.SourceTable;
                                xlWorkSheet.Cells[10][linha] = rid.Indicators;
                                linha++;
                                }
                            if (!set.Equals(datim.FILTRO_VALOR_DEFEITO))
                            {if (rid.Value == null)
                                { rid.Value = "0"; }
                                if (linha == 1)
                                {
                                    xlApp = new Excel.Application();
                                    xlApp.Visible = false;
                                    misValue = System.Reflection.Missing.Value;
                                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                                    xlWorkSheet.Cells[1][linha] = "ID_DATIM";
                                    xlWorkSheet.Cells[2][linha] = "IndicatorMapping.Desag";
                                    xlWorkSheet.Cells[3][linha] = "Support_Type";
                                    xlWorkSheet.Cells[4][linha] = "IndicatorMapping.EnteryFildID";
                                    xlWorkSheet.Cells[5][linha] = "Value";
                                    xlWorkSheet.Cells[6][linha] = "DataSet";
                                    xlWorkSheet.Cells[7][linha] = "Distrito_DATIM";
                                    xlWorkSheet.Cells[8][linha] = "Nome_US_DATIM";
                                    xlWorkSheet.Cells[9][linha] = "Source.Table";
                                    xlWorkSheet.Cells[10][linha] = "Indicators";
                                    linha++;
                                }
                                xlWorkSheet.Cells[1][linha] = rid.Id;
                                xlWorkSheet.Cells[2][linha] = rid.Desagregado;
                                xlWorkSheet.Cells[3][linha] = rid.SuportType;
                                xlWorkSheet.Cells[4][linha] = rid.EntryField;
                                xlWorkSheet.Cells[5][linha] = rid.Value;
                                xlWorkSheet.Cells[6][linha] = rid.DataSet;
                                xlWorkSheet.Cells[7][linha] = rid.Distrito;
                                xlWorkSheet.Cells[8][linha] = rid.US;
                                xlWorkSheet.Cells[9][linha] = rid.SourceTable;
                                xlWorkSheet.Cells[10][linha] = rid.Indicators;
                                linha++;
                            }
                        }
                        if(linha>=2)
                        {
                                xlWorkSheet.Columns.AutoFit();
                            xlWorkBook.WebOptions.Encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;
                            xlWorkBook.SaveAs("DATIM_"+i+"_"+u , Excel.XlFileFormat.xlWorkbookNormal ,      misValue , misValue , misValue , misValue , Excel.XlSaveAsAccessMode.xlExclusive , misValue , misValue , misValue , misValue , misValue);
                                xlWorkBook.Close(true , misValue , misValue);
                                xlApp.Quit();
                          
                                Marshal.ReleaseComObject(xlWorkSheet);
                                Marshal.ReleaseComObject(xlWorkBook);
                                Marshal.ReleaseComObject(xlApp);
                            
                        }
                      
                        progress++;
                        Dispatcher.Invoke(() =>
                        {
                           
                            PBar.Value = (progress * 100) / (fi.Count() * fus.Count());
                            PBarText.Text = Math.Round(PBar.Value, 2) + "%";
                        });


                    }


                }
                Dispatcher.Invoke(() =>
                {
                    MessageBox.Show("Ficheiros gerados com sucesso", "Gerador", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                botaoGerar.IsEnabled = true;
                comboIndicadorListaCampo.IsEnabled = true;
                comboIndicadorListaFiltros.IsEnabled = true;
                comboUSListaCampo.IsEnabled = true;
                comboUSListaFiltros.IsEnabled = true;
                comboValueListaCampo.IsEnabled = true;
                    comboValueListaFiltros.IsEnabled = true;
                    botaoSelecionar.IsEnabled = true;
                    PBar.Visibility = Visibility.Hidden;
                    PBarText.Visibility = Visibility.Hidden;
                });
            });
            


            }
        private void botaoSair(object sender , RoutedEventArgs e) { }
        }

    }