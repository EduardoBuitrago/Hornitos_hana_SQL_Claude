using Lotes_Hornitos.Modelos;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Linq;
using System.Security.Cryptography;

namespace Lotes_Hornitos.Classes
{
    internal class FormLotesSalida
    {
        private SAPbobsCOM.Company oCompany;
        private SAPbouiCOM.Application oApp;
        public static List<LoteSalida> listaLotesSalida;
        public int SelectionManual = 0;

        /// <summary>
        /// Constructor de la clase
        /// </summary>
        /// <param name="oCompany"></param>
        /// <param name="oApp"></param>
        public FormLotesSalida(SAPbobsCOM.Company oCompany, SAPbouiCOM.Application oApp)
        {
            this.oCompany = oCompany;
            this.oApp = oApp;
            listaLotesSalida = new List<LoteSalida>();
        }


        /// <summary>
        /// Funcion para agregar informacion de lote para salida
        /// </summary>
        /// <param name="pVal"></param>
        internal void AddInfoLotes(ItemEvent pVal)
        {
            SAPbouiCOM.Form Salida = null;
            Grid grid_Disponible = null;
            Grid grid_LineasDoc = null;
            Grid grid_Seleccion = null;
            DataTable DT_Seleccion = null;
            DataTable DT_Disponible = null;
            DataTable DT_LineasDocumento = null;
            EditTextColumn ColumnGrid = null;

            int cantidad = 0;

            string query = string.Empty;

            bool EnableGridDisponible = true;

            var Sb = new System.Text.StringBuilder();

            try
            {
                Salida = oApp.Forms.Item(pVal.FormUID);
                Salida.Freeze(true);

                grid_Disponible = (Grid)Salida.Items.Item("GrdDisp").Specific;
                grid_LineasDoc = (Grid)Salida.Items.Item("GrdLineas").Specific;
                grid_Seleccion = (Grid)Salida.Items.Item("GrdSel").Specific;
                DT_Disponible = (DataTable)Salida.DataSources.DataTables.Item("DTDispo");
                DT_LineasDocumento = (DataTable)Salida.DataSources.DataTables.Item("DTLineas");
                DT_Seleccion = (DataTable)Salida.DataSources.DataTables.Item("DTSel");

                if (SelectionManual.Equals(1))
                {
                    //cantidad = (int)DT_LineasDocumento.GetValue("Cantidad", pVal.Row);

                    //Valida que las cantidades entre cantidad de linea de docuemnto sea igal a la cantidad de lineas en seelccionados
                    EnableGridDisponible = ValidarCantidadesDisponibleToSeleccionados(DT_LineasDocumento.GetValue("ItemCode", pVal.Row).ToString(), Convert.ToDouble(DT_LineasDocumento.GetValue("Cantidad", pVal.Row)));

                    query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.InfoArticulosconLoteUDT_Salida_Manual.sql")).ReadToEnd();
                    query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, DT_LineasDocumento.GetValue("ItemCode", pVal.Row), DT_LineasDocumento.GetValue("Almacen", pVal.Row)).ToString();

                    DT_Disponible.ExecuteQuery(query);
                    DT_Seleccion.Clear();



                    if (!DT_Disponible.IsEmpty)
                    {
                        grid_Disponible.SelectionMode = BoMatrixSelect.ms_None;
                        grid_Disponible.Item.Enabled = true;
                        grid_Disponible.Item.AffectsFormMode = true;
                        grid_Disponible.Columns.Item("ItemCode").Visible = false;
                        grid_Disponible.Columns.Item("Tipo doc").Visible = false;
                        grid_Disponible.Columns.Item("Num Doc").Visible = false;
                        grid_Disponible.Columns.Item("Almacen").Visible = false;

                        ColumnGrid = (EditTextColumn)grid_Disponible.Columns.Item("Sel");
                        ColumnGrid.Type = BoGridColumnType.gct_CheckBox;


                        for (int i = 0; i < grid_Disponible.Columns.Count; i++)
                        {

                            grid_Disponible.Columns.Item(i).AffectsFormMode = false;
                            grid_Disponible.Columns.Item(i).TitleObject.Sortable = false;
                            grid_Disponible.Columns.Item(i).Editable = false;
                        }

                        grid_Disponible.Columns.Item("Cantidad").Editable = true;
                        grid_Disponible.Columns.Item("Sel").Editable = true;
                        grid_Disponible.Columns.Item("Cantidad").AffectsFormMode = true;
                        grid_Disponible.Columns.Item("Fecha vence").Editable = false;
                        grid_Disponible.Columns.Item("Fecha Admin").Editable = false;

                        grid_Disponible.Columns.Item("Fecha vence").TitleObject.Sortable = true;
                        grid_Disponible.Columns.Item("Fecha vence").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
                        //grid_Disponible.Columns.Item("Cantidad").AffectsFormMode = true;

                        

                        grid_Disponible.AutoResizeColumns();

                        grid_Disponible.Item.Enabled = EnableGridDisponible;

                    }
                    else
                    {
                        grid_Disponible.SelectionMode = BoMatrixSelect.ms_None;
                        grid_Disponible.Item.Enabled = false;
                    }
                }


                infolotesentablasalida(ref DT_Seleccion, ref grid_Seleccion, DT_LineasDocumento.GetValue("ItemCode", pVal.Row).ToString(), DT_LineasDocumento.GetValue("Almacen", pVal.Row).ToString(), (int)DT_LineasDocumento.GetValue("NumLinea", pVal.Row));

                for (int i = 0; i < DT_Disponible.Rows.Count; i++)
                {
                    for (int j = 0; j < listaLotesSalida.Count; j++)
                    {
                        var item = listaLotesSalida[j];
                        var itemcode = DT_Disponible.GetValue("ItemCode", i).ToString();
                            if (item.ItemCode.Equals(DT_Disponible.GetValue("ItemCode", i).ToString()))
                            {
                                if (item.Almacen.Equals(DT_Disponible.GetValue("Almacen", i).ToString()))
                                {
                                    if (item.NumeroLote.Equals(DT_Disponible.GetValue("Numero lote", i).ToString()))
                                    {
                                        DT_Disponible.SetValue("Sel", i, "Y");
                                        DT_Disponible.SetValue("Cantidad", i, item.Cantidad);
                                    }                                        
                                }
                            }
                        

                    }
                }



                Salida.Freeze(false);
            }
            catch (Exception ex)
            {
                Salida.Freeze(false);
                oApp.SetStatusBarMessage("Error Add-On " + Ini.asm.GetName().Name + " : " + ex.Message.ToString() + " " + ex.StackTrace.ToString());

            }
            finally
            {
                if (Salida != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Salida);
                }
                if (grid_Disponible != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(grid_Disponible);
                }
                if (grid_LineasDoc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(grid_LineasDoc);
                }
                if (grid_Seleccion != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(grid_Seleccion);
                }
                if (DT_Seleccion != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DT_Seleccion);
                }
                if (DT_LineasDocumento != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DT_LineasDocumento);
                }
                if (DT_Disponible != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DT_Disponible);
                }
                if (ColumnGrid != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ColumnGrid);
                }
            }
        }

        /// <summary>
        /// Valida que las cantidades entre cantidad de linea de docuemnto sea igual a la cantidad de lineas en seelccionados
        /// </summary>
        /// <param name="oItem">Item a evaluar las cantidades</param>
        /// <returns></returns>
        private bool ValidarCantidadesDisponibleToSeleccionados(string oItem, double CantidadLinea)
        {
            double cantidadSeleccionada = 0;
            bool result = false;

            if (listaLotesSalida == null || !listaLotesSalida.Any())
            {
                //DesdeLista = false;
                result = true;
                return result;
            }



            for (int j = 0; j < listaLotesSalida.Count; j++)
            {
                var item = listaLotesSalida[j];

                if (item.ItemCode.Equals(oItem))
                {
                    cantidadSeleccionada += Convert.ToDouble(item.Cantidad);

                }
            }




            if (!cantidadSeleccionada.Equals(CantidadLinea))
            {
                result = true;
            }
            return result;
        }

        private void infolotesentablasalida(ref DataTable dT_Seleccion, ref Grid grid_Seleccion, string ItemCode, string almacen, int NumLinea)
        {

            CargarLista(ref dT_Seleccion, listaLotesSalida, ItemCode, almacen, NumLinea);


            //var Sb = new System.Text.StringBuilder();
            //string query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.InfoArticulosUDTSalida.sql")).ReadToEnd();
            //query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, ItemCode, NumDoc, TipoDoc).ToString();


            //dT_Seleccion.ExecuteQuery(query);


            if (!dT_Seleccion.IsEmpty)
            {
                grid_Seleccion.SelectionMode = BoMatrixSelect.ms_None;
                grid_Seleccion.Item.Enabled = false;
                grid_Seleccion.Item.AffectsFormMode = true;
                grid_Seleccion.Columns.Item("ItemCode").Visible = false;
                grid_Seleccion.Columns.Item("Sel").Visible = false;
                grid_Seleccion.Columns.Item("Tipo doc").Visible = false;
                grid_Seleccion.Columns.Item("Num Doc").Visible = false;
                grid_Seleccion.Columns.Item("NumLinea").Visible = false;



                for (int i = 0; i < grid_Seleccion.Columns.Count; i++)
                {

                    grid_Seleccion.Columns.Item(i).AffectsFormMode = false;
                    grid_Seleccion.Columns.Item(i).TitleObject.Sortable = false;
                    grid_Seleccion.Columns.Item(i).Editable = false;
                }
                grid_Seleccion.SelectionMode = BoMatrixSelect.ms_Single;
                grid_Seleccion.AutoResizeColumns();
            }
            else
            {
                grid_Seleccion.SelectionMode = BoMatrixSelect.ms_None;
                grid_Seleccion.Item.Enabled = false;
            }
        }


        private void CargarLista(ref DataTable dT_Seleccion, List<LoteSalida> listaLotesSalida, string _ItemCode, string almacen, int numLinea)
        {
            string query = string.Empty;
            int cont = 0;
            List<string> ListLotes = new List<string>();
            var Sb = new System.Text.StringBuilder();

            dT_Seleccion.Rows.Clear();

            if (listaLotesSalida == null || !listaLotesSalida.Any())
            {
                //DesdeLista = false;
                return;
            }


            query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.LotesDatatable_salida.sql")).ReadToEnd();


            for (int i = 0; i < listaLotesSalida.Count; i++)
            {
                var item = listaLotesSalida[i];
                if (item.NumLinea.Equals(numLinea))
                {
                    if (item.ItemCode.Equals(_ItemCode))
                    {
                        if (item.Almacen.Equals(almacen))
                        {
                            ListLotes.Add(Sb.AppendFormat(CultureInfo.InvariantCulture, query, item.ItemCode, item.NumeroDocumento, item.TipoDocumento, item.Cantidad, item.NumeroLote, item.FechaProceso, item.Almacen, i, item.NumLinea).ToString());
                            Sb.Clear();
                        }
                    }
                }

            }

            if (ListLotes.Any())
            {
                query = string.Join(" UNION ALL ", ListLotes);
                dT_Seleccion.ExecuteQuery(query);
                //DesdeLista = true;
            }
            //else
            //{
            //    DesdeLista = false;
            //}
        }


        /// <summary>
        /// Funciona que limpia los check selecionadso y el datable de lotes disponibles.
        /// </summary>
        /// <param name="pVal"></param>
        internal void CleanData(ItemEvent pVal)
        {
            SAPbouiCOM.Form Salida = null;
            Grid grid_Disponible = null;
            Grid grid_LineasDoc = null;
            DataTable DT_Disponible = null;
            DataTable DT_LineasDocumento = null;

            try
            {

                Salida = oApp.Forms.Item(pVal.FormUID);
                Salida.Freeze(true);
                grid_Disponible = (Grid)Salida.Items.Item("GrdDisp").Specific;
                grid_LineasDoc = (Grid)Salida.Items.Item("GrdLineas").Specific;
                DT_Disponible = (DataTable)Salida.DataSources.DataTables.Item("DTDispo");
                DT_LineasDocumento = (DataTable)Salida.DataSources.DataTables.Item("DTLineas");

                for (int i = 0; i < DT_LineasDocumento.Rows.Count; i++)
                {
                    if (DT_LineasDocumento.GetValue("Sel", i).Equals("Y"))
                    {
                        DT_LineasDocumento.SetValue("Sel", i, "N");
                    }
                }

                DT_Disponible.Clear();

                Salida.Freeze(false);

            }
            catch (Exception ex)
            {
                Salida.Freeze(false);

            }
            finally
            {
                if (Salida != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Salida);
                }
                if (grid_Disponible != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(grid_Disponible);
                }
                if (grid_LineasDoc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(grid_LineasDoc);
                }
                if (DT_LineasDocumento != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DT_LineasDocumento);
                }
                if (DT_Disponible != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DT_Disponible);
                }
            }
        }


        /// <summary>
        /// Funcion para mostrar el formulario de salida de lotes
        /// </summary>
        /// <param name="query"></param>
        /// <returns></returns>
        private bool showForm(string query)
        {
            System.Xml.XmlDocument oXmlDocument = null;

            string stringForm = null, oPath = null, FormUID = null;

            SAPbouiCOM.Form oForm = null;
            Grid grid_LineasDoc = null;
            DataTable DT_LineasDocumento = null;
            EditTextColumn ColumnGrid = null;

            List<bool> result = new List<bool>();
            bool res = true, existeenconsulta = false;
            try
            {
                System.Console.WriteLine("Inico carga de formulario XML FRMLTS.srf de salida de lotes");
                oXmlDocument = new System.Xml.XmlDocument();

                oPath = $"{Assembly.GetExecutingAssembly().GetName().Name}.Resources.XML.FRMLTS.srf";


                Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(oPath);
                StreamReader sr = new StreamReader(stream);
                stringForm = sr.ReadToEnd();

                oXmlDocument.LoadXml(stringForm);
                FormUID = oXmlDocument.SelectSingleNode("/application/forms/action/form | /Application/forms/action/form").Attributes["uid"].Value;
                // CORRECCION: si la ventana ya estaba abierta (por ejemplo el
                // usuario vuelve a pedir el formulario sin haberlo cerrado), se
                // cierra primero para poder abrirla de nuevo limpia. Antes esto
                // lanzaba "Form - already exists [66000-11]".
                try
                {
                    SAPbouiCOM.Form oFormExistente = oApp.Forms.Item(FormUID);
                    oFormExistente.Close();
                    Ini.ReleaseCOMObjetc(oFormExistente);
                }
                catch
                {
                    // No existe todavía, es el caso normal. No hay nada que hacer.
                }

                oApp.LoadBatchActions(ref stringForm);

                System.Console.WriteLine("Ajuste de campos de formulario de salida de lotes");

                oForm = oApp.Forms.Item(FormUID);

                grid_LineasDoc = (Grid)oForm.Items.Item("GrdLineas").Specific;

                DT_LineasDocumento = (DataTable)oForm.DataSources.DataTables.Item("DTLineas");

                //validar si ya se agregaron los lotes
                result = ValidasiexisteloteentablaSalida(query);


                System.Console.WriteLine("Ejuctar query de formulario de salida de lotes");

                DT_LineasDocumento.ExecuteQuery(query);


                if (!DT_LineasDocumento.IsEmpty)
                {

                    foreach (bool l in result)
                    {
                        if (!l)
                        {
                            res = l;
                            break;
                        }
                    }
                    if (!res)
                    {
                        System.Console.WriteLine("Formateo de campos de formulario de salida de lotes");
                        grid_LineasDoc.SelectionMode = BoMatrixSelect.ms_None;
                        grid_LineasDoc.Item.Enabled = true;
                        grid_LineasDoc.Item.AffectsFormMode = false;


                        ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("ItemCode");
                        ColumnGrid.LinkedObjectType = "4";
                        ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("Almacen");
                        ColumnGrid.LinkedObjectType = "64";
                        if (Ini.Tipo.Equals("OWTR"))
                        {
                            ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("Almacen destino");
                            ColumnGrid.LinkedObjectType = "64";
                        }
                        ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("Sel");
                        ColumnGrid.Type = BoGridColumnType.gct_CheckBox;


                        for (int i = 0; i < grid_LineasDoc.Columns.Count; i++)
                        {

                            grid_LineasDoc.Columns.Item(i).Editable = false;
                            grid_LineasDoc.Columns.Item(i).TitleObject.Sortable = false;
                            grid_LineasDoc.Columns.Item(i).AffectsFormMode = false;
                        }
                        if (SelectionManual.Equals(1))
                        {
                            grid_LineasDoc.Columns.Item("Sel").Editable = true;
                            oForm.Items.Item("BtnAutSel").Enabled = false;
                        }
                        else
                        {
                            grid_LineasDoc.Columns.Item("Sel").Editable = false;
                        }


                        grid_LineasDoc.Columns.Item("ItemCode").TitleObject.Caption = "Número de articulo";
                        //grid_LineasDoc.Columns.Item("Cantidad solicitada").TitleObject.Caption = "Cantidad";
                        grid_LineasDoc.Columns.Item("Numero Doc").TitleObject.Caption = "Número documento";
                        grid_LineasDoc.Columns.Item("Tipo Doc").Visible = false;
                        grid_LineasDoc.AutoResizeColumns();
                        System.Console.WriteLine("Mostrar formulario de salida de lotes");
                        oForm.Visible = true;
                    }
                    else
                    {
                        oForm.Close();
                    }

                }
                else
                {
                    grid_LineasDoc.SelectionMode = BoMatrixSelect.ms_None;
                    grid_LineasDoc.Item.Enabled = false;
                    System.Console.WriteLine("Cerrar formulario de salida de lotes");
                    oForm.Close();
                    res = true;

                }


                return res;

            }
            catch (Exception ex)
            {
                oApp.SetStatusBarMessage("Error Add-On " + Ini.asm.GetName().Name + " : " + ex.Message.ToString() + " " + ex.StackTrace.ToString());
                return false;

            }
            finally
            {
                if (oForm != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                }
                if (grid_LineasDoc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(grid_LineasDoc);
                }
                if (DT_LineasDocumento != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DT_LineasDocumento);
                }
                if (ColumnGrid != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ColumnGrid);
                }
            }
        }


        /// <summary>
        /// Funcion que valida que el lote exista en la tabla temporal
        /// </summary>
        /// <param name="query">Consulta con la informacion de los lotes</param>
        /// <returns></returns>
        private List<bool> ValidasiexisteloteentablaSalida(string query)
        {
            double cantidadLote = 0;
            bool existelote = false, existeenconsulta = false;
            SAPbobsCOM.Recordset oRs = null;

            List<bool> result = new List<bool>();

            try
            {

                if (listaLotesSalida == null || !listaLotesSalida.Any())
                {
                    result.Add(false);
                    return result;
                }



                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRs.DoQuery(query);
                oRs.MoveFirst();

                for (int i = 1; i <= oRs.RecordCount; i++)
                {

                    existelote = false;
                    cantidadLote = 0;
                    for (int j = 0; j < listaLotesSalida.Count; j++)
                    {
                        var item = listaLotesSalida[j];

                        System.Console.WriteLine($"desde lista {item.ItemCode}");
                        System.Console.WriteLine($"desde recordset {oRs.Fields.Item("ItemCode").Value.ToString()}");

                        if (item.NumLinea.Equals((int)oRs.Fields.Item("NumLinea").Value))
                        {
                            if (item.ItemCode.Equals(oRs.Fields.Item("ItemCode").Value.ToString()))
                            {

                                if (!item.Almacen.Equals(oRs.Fields.Item("Almacen").Value.ToString()))
                                {
                                    listaLotesSalida.RemoveAt(j);
                                    result.Add(false);
                                    break;
                                }

                                cantidadLote += item.Cantidad;
                                existelote = true;
                            }
                        }
                    }
                    if (existelote)
                    {
                        if (cantidadLote < Convert.ToDouble(oRs.Fields.Item("Cantidad").Value))
                        {
                            result.Add(false);
                            break;
                        }
                    }
                    else
                    {
                        result.Add(false);
                    }
                    oRs.MoveNext();
                }



                //validar que los valores de la tabla tengan los mismos de la consulta.
                int o = 0;
                for (o = 0; o < listaLotesSalida.Count; o++)
                {
                    existeenconsulta = false;
                    var item = listaLotesSalida[o];
                    oRs.MoveFirst();
                    for (int i = 1; i <= oRs.RecordCount; i++)
                    {
                        if (item.NumLinea.Equals((int)oRs.Fields.Item("NumLinea").Value))
                        {
                            if (item.ItemCode.Equals(oRs.Fields.Item("ItemCode").Value.ToString()))
                            {
                                if (item.Almacen.Equals(oRs.Fields.Item("Almacen").Value.ToString()))
                                {
                                    existeenconsulta = true;
                                }
                            }
                        }


                        oRs.MoveNext();
                    }
                    if (!existeenconsulta)
                    {
                        listaLotesSalida.Remove(item);
                        o--;
                    }

                }

                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en ValidasiexisteloteentablSalida {ex.Message}");
                return result;
            }
            finally
            {
                if (oRs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
                }
            }
        }

        //private List<bool> ValidasiexisteloteentablaSalida(string query, string oDocNum, int oCantidad, string almacen)
        //{

        //    int cantidadLote = 0;
        //    bool existelote = false;
        //    try
        //    {

        //        if (listaLotesSalida == null || !listaLotesSalida.Any())
        //            return false;

        //        for (int i = 0; i < listaLotesSalida.Count; i++)
        //        {
        //            var item = listaLotesSalida[i];
        //            if (item.ItemCode.Equals(_ItemCode))
        //            {
        //                if (!item.Almacen.Equals(almacen))
        //                {
        //                    listaLotesSalida.RemoveAt(i);
        //                    return false;                            
        //                }

        //                cantidadLote += item.Cantidad;
        //                existelote = true;
        //            }


        //        }

        //        if (existelote)
        //        {
        //            if (cantidadLote < oCantidad)
        //            {
        //                return false;
        //            }
        //        }
        //        else
        //        {
        //            return false;
        //        }
        //        //oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //        //query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ArticulosenTablaSalida.sql")).ReadToEnd();
        //        //query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, ItemCode, tipo, oDocNum).ToString();
        //        //oRs.DoQuery(query);
        //        //if (oRs == null)
        //        //{
        //        //    return false;
        //        //}
        //        //else
        //        //{
        //        //    oRs.MoveFirst();

        //        //    int valor = (int)oRs.Fields.Item(0).Value;
        //        //    if (valor.Equals(0)) return false;
        //        //}


        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"Error en ValidasiexisteloteentablaSalida {ex.Message}");
        //        return false;
        //    }
        //    finally
        //    {
        //        //Ini.ReleaseCOMObjetc(oRs);
        //    }
        //}


        /// <summary>
        /// Mostrar los items de la salida de lotes
        /// </summary>
        /// <param name="pVal"></param>
        internal bool MostarItems(ItemEvent pVal, string Tipo)
        {
            string query = string.Empty;
            string path = string.Empty;

            SAPbouiCOM.Form Salida = null;
            SAPbouiCOM.Matrix LineasEntrada = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.EditText oCantidad = null;
            SAPbouiCOM.EditText oAlmacen = null;
            SAPbouiCOM.EditText oAlmacenDestino = null;
            SAPbouiCOM.EditText oDocNum = null;
            SAPbouiCOM.EditText RowNumber = null;
            SAPbouiCOM.EditText BaseType = null;

            double cantidad = 0;
            List<string> Items = new List<string>();
            var Sb = new System.Text.StringBuilder();
            bool result = false;
            try
            {
                Console.WriteLine("Inicio de consulta de articulos con lotes salida");

                Salida = oApp.Forms.Item(pVal.FormUID);

                if (!Salida.Mode.Equals(BoFormMode.fm_ADD_MODE)) return true;

                switch (Tipo)
                {
                    case "OIGE":
                        LineasEntrada = (SAPbouiCOM.Matrix)Salida.Items.Item("13").Specific;
                        query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ArticulosconLoteSalida.sql")).ReadToEnd();
                        break;
                    case "OWTR":
                        LineasEntrada = (SAPbouiCOM.Matrix)Salida.Items.Item("23").Specific;
                        query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ArticulosconLoteSalidaTransferencia.sql")).ReadToEnd();
                        break;
                    default:
                        LineasEntrada = (SAPbouiCOM.Matrix)Salida.Items.Item("38").Specific;
                        query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ArticulosconLoteSalida.sql")).ReadToEnd();
                        break;

                }
                //dBDataSourceLineasDocumento = Entrada.DataSources.DBDataSources.Item("PDN1");



                Console.WriteLine("Armado de query de articulos con lotes salida");

                for (int i = 1; i <= LineasEntrada.RowCount; i++)
                {
                    oEditText = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("1").Cells.Item(i).Specific;

                    if (!string.IsNullOrEmpty(oEditText.String))
                    {
                        switch (Tipo)
                        {
                            case "OINV":
                                BaseType = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("43").Cells.Item(i).Specific;
                                if (BaseType.Value.Equals("15"))
                                {
                                    return true;
                                }
                                break;
                            case "ORPC":
                                BaseType = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("43").Cells.Item(i).Specific;
                                if (BaseType.Value.Equals("21"))
                                {
                                    return true;
                                }
                                break;
                        }

                        switch (Tipo)
                        {
                            case "OIGE":
                                oCantidad = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("10001028").Cells.Item(i).Specific;
                                break;
                            case "OWTR":
                                oCantidad = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("10001029").Cells.Item(i).Specific;
                                break;
                            default:
                                oCantidad = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("10002117").Cells.Item(i).Specific;
                                break;

                        }

                        switch (Tipo)
                        {
                            case "OIGE":
                                oAlmacen = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("15").Cells.Item(i).Specific;
                                break;
                            case "OWTR":
                                oAlmacen = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("1470001039").Cells.Item(i).Specific;
                                oAlmacenDestino = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("5").Cells.Item(i).Specific;
                                break;
                            default:
                                oAlmacen = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("24").Cells.Item(i).Specific;

                                break;

                        }

                        switch (Tipo)
                        {
                            case "OIGE":
                                oDocNum = (SAPbouiCOM.EditText)Salida.Items.Item("7").Specific;
                                break;
                            case "OWTR":
                                oDocNum = (SAPbouiCOM.EditText)Salida.Items.Item("11").Specific;
                                break;
                            default:
                                oDocNum = (SAPbouiCOM.EditText)Salida.Items.Item("8").Specific;
                                break;

                        }

                        RowNumber = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("0").Cells.Item(i).Specific;
                        cantidad = Ini.ConvertToSystemFormat(oCantidad.String);

                        switch (Tipo)
                        {
                            case "OWTR":
                                Items.Add(Sb.AppendFormat(CultureInfo.InvariantCulture, query, oEditText.String, oDocNum.String, oAlmacen.String, cantidad, Tipo, oAlmacenDestino.String, RowNumber.Value).ToString());
                                break;
                            default:
                                Items.Add(Sb.AppendFormat(CultureInfo.InvariantCulture, query, oEditText.String, oDocNum.String, oAlmacen.String, cantidad, Tipo, RowNumber.Value).ToString());
                                break;
                        }
                        Sb.Clear();
                        System.Console.WriteLine(oEditText.String);

                    }
                }

                //query = string.Join(",", Items.Select(x => $"'{x}'"));
                query = string.Join(" UNION ALL ", Items);


                Console.WriteLine("fin Armado de query de articulos con lotes salida");


                // Mostar el formulario de lotes.
                System.Console.WriteLine("Mostrar el formulario de salida de lotes");
                result = showForm(query);
                System.Console.WriteLine("Fin carga de formulario de salida de lotes");

                return result;

            }
            catch (Exception ex)
            {
                Salida.Freeze(false);
                oApp.SetStatusBarMessage("Error Add-On " + Ini.asm.GetName().Name + " : " + ex.Message.ToString() + " " + ex.StackTrace.ToString());
                return false;
            }
            finally
            {
                if (Salida != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Salida);
                }
                if (LineasEntrada != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(LineasEntrada);
                }
                if (oEditText != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEditText);
                }
                if (oCantidad != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCantidad);
                }
                if (oAlmacen != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oAlmacen);
                }
                if (oAlmacenDestino != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oAlmacenDestino);
                }
                if (oDocNum != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDocNum);
                }
                if (RowNumber != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(RowNumber);
                }
            }
        }

        /// <summary>
        /// Funcion para mostar los lotes de salida
        /// </summary>
        /// <param name="oForm">Formulario activo</param>
        /// <param name="Tipo">Tipo de tabla de base de datos</param>
        /// <returns></returns>
        internal void MostarItems(SAPbouiCOM.Form oForm, string Tipo)
        {
            string query = string.Empty;
            string path = string.Empty;

            SAPbouiCOM.Form Salida = null;
            SAPbouiCOM.Matrix LineasEntrada = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.EditText oCantidad = null;
            SAPbouiCOM.EditText oAlmacen = null;
            SAPbouiCOM.EditText oAlmacenDestino = null;
            SAPbouiCOM.EditText oDocNum = null;
            SAPbouiCOM.EditText RowNumber = null;
            SAPbouiCOM.EditText BaseType = null;

            double cantidad = 0;

            List<string> Items = new List<string>();
            var Sb = new System.Text.StringBuilder();
            bool result = false;
            try
            {
                Console.WriteLine("Inicio de consulta de articulos con lotes salida");

                Salida = oForm;

                if (!Salida.Mode.Equals(BoFormMode.fm_ADD_MODE)) return;

                switch (Tipo)
                {
                    case "OIGE":
                        LineasEntrada = (SAPbouiCOM.Matrix)Salida.Items.Item("13").Specific;
                        query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ArticulosconLoteSalida.sql")).ReadToEnd();
                        break;
                    case "OWTR":
                        LineasEntrada = (SAPbouiCOM.Matrix)Salida.Items.Item("23").Specific;
                        query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ArticulosconLoteSalidaTransferencia.sql")).ReadToEnd();
                        break;
                    default:
                        LineasEntrada = (SAPbouiCOM.Matrix)Salida.Items.Item("38").Specific;
                        query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ArticulosconLoteSalida.sql")).ReadToEnd();
                        break;

                }
                //dBDataSourceLineasDocumento = Entrada.DataSources.DBDataSources.Item("PDN1");



                Console.WriteLine("Armado de query de articulos con lotes salida");

                for (int i = 1; i <= LineasEntrada.RowCount; i++)
                {
                    oEditText = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("1").Cells.Item(i).Specific;

                    if (!string.IsNullOrEmpty(oEditText.String))
                    {
                        //switch (Tipo)
                        //{
                        //case "OINV":
                        //    BaseType = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("43").Cells.Item(i).Specific;
                        //    if (BaseType.Value.Equals("15"))
                        //    {
                        //        return true;
                        //    }
                        //    break;
                        //case "ORPC":
                        //    BaseType = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("43").Cells.Item(i).Specific;
                        //    if (BaseType.Value.Equals("21"))
                        //    {
                        //        return true;
                        //    }
                        //    break;
                        //}

                        switch (Tipo)
                        {
                            case "OIGE":
                                oCantidad = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("10001028").Cells.Item(i).Specific;
                                break;
                            case "OWTR":
                                oCantidad = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("10001029").Cells.Item(i).Specific;
                                break;
                            default:
                                oCantidad = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("10002117").Cells.Item(i).Specific;
                                break;

                        }

                        switch (Tipo)
                        {
                            case "OIGE":
                                oAlmacen = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("15").Cells.Item(i).Specific;
                                break;
                            case "OWTR":
                                oAlmacen = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("1470001039").Cells.Item(i).Specific;
                                oAlmacenDestino = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("5").Cells.Item(i).Specific;
                                break;
                            default:
                                oAlmacen = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("24").Cells.Item(i).Specific;

                                break;

                        }

                        switch (Tipo)
                        {
                            case "OIGE":
                                oDocNum = (SAPbouiCOM.EditText)Salida.Items.Item("7").Specific;
                                break;
                            case "OWTR":
                                oDocNum = (SAPbouiCOM.EditText)Salida.Items.Item("11").Specific;
                                break;
                            default:
                                oDocNum = (SAPbouiCOM.EditText)Salida.Items.Item("8").Specific;
                                break;

                        }

                        RowNumber = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("0").Cells.Item(i).Specific;
                        cantidad = Ini.ConvertToSystemFormat(oCantidad.String);

                        switch (Tipo)
                        {
                            case "OWTR":
                                Items.Add(Sb.AppendFormat(CultureInfo.InvariantCulture, query, oEditText.String, oDocNum.String, oAlmacen.String, cantidad, Tipo, oAlmacenDestino.String, RowNumber.Value).ToString());
                                break;
                            default:
                                Items.Add(Sb.AppendFormat(CultureInfo.InvariantCulture, query, oEditText.String, oDocNum.String, oAlmacen.String, cantidad, Tipo, RowNumber.Value).ToString());
                                break;
                        }
                        Sb.Clear();
                        System.Console.WriteLine(oEditText.String);

                    }
                }

                //query = string.Join(",", Items.Select(x => $"'{x}'"));
                query = string.Join(" UNION ALL ", Items);


                Console.WriteLine("fin Armado de query de articulos con lotes salida");


                // Mostar el formulario de lotes.
                System.Console.WriteLine("Mostrar el formulario de salida de lotes");
                showFormMostrarLotes(query);
                System.Console.WriteLine("Fin carga de formulario de salida de lotes");

                //return result;

            }
            catch (Exception ex)
            {
                Salida.Freeze(false);
                oApp.SetStatusBarMessage("Error Add-On " + Ini.asm.GetName().Name + " : " + ex.Message.ToString() + " " + ex.StackTrace.ToString());
                //return false;
            }
            finally
            {
                if (Salida != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Salida);
                }
                if (LineasEntrada != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(LineasEntrada);
                }
                if (oEditText != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEditText);
                }
                if (oCantidad != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCantidad);
                }
                if (oAlmacen != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oAlmacen);
                }
                if (oAlmacenDestino != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oAlmacenDestino);
                }
                if (oDocNum != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDocNum);
                }
                if (RowNumber != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(RowNumber);
                }
            }
        }

        /// <summary>
        /// Mostar formulariode lotes sin condicionales
        /// </summary>
        /// <param name="query">Consulta con informacion de lotes a mostrar</param>
        /// <returns></returns>
        private void showFormMostrarLotes(string query)
        {
            System.Xml.XmlDocument oXmlDocument = null;

            string stringForm = null, oPath = null, FormUID = null;

            SAPbouiCOM.Form oForm = null;
            Grid grid_LineasDoc = null;
            DataTable DT_LineasDocumento = null;
            EditTextColumn ColumnGrid = null;

            //List<bool> result = new List<bool>();
            //bool res = true, existeenconsulta = false;
            try
            {
                System.Console.WriteLine("Inico carga de formulario XML FRMLTS.srf de salida de lotes");
                oXmlDocument = new System.Xml.XmlDocument();

                oPath = $"{Assembly.GetExecutingAssembly().GetName().Name}.Resources.XML.FRMLTS.srf";


                Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(oPath);
                StreamReader sr = new StreamReader(stream);
                stringForm = sr.ReadToEnd();

                oXmlDocument.LoadXml(stringForm);
                FormUID = oXmlDocument.SelectSingleNode("/application/forms/action/form | /Application/forms/action/form").Attributes["uid"].Value;
                // CORRECCION: si la ventana ya estaba abierta (por ejemplo el
                // usuario vuelve a pedir el formulario sin haberlo cerrado), se
                // cierra primero para poder abrirla de nuevo limpia. Antes esto
                // lanzaba "Form - already exists [66000-11]".
                try
                {
                    SAPbouiCOM.Form oFormExistente = oApp.Forms.Item(FormUID);
                    oFormExistente.Close();
                    Ini.ReleaseCOMObjetc(oFormExistente);
                }
                catch
                {
                    // No existe todavía, es el caso normal. No hay nada que hacer.
                }

                oApp.LoadBatchActions(ref stringForm);

                System.Console.WriteLine("Ajuste de campos de formulario de salida de lotes");

                oForm = oApp.Forms.Item(FormUID);

                grid_LineasDoc = (Grid)oForm.Items.Item("GrdLineas").Specific;

                DT_LineasDocumento = (DataTable)oForm.DataSources.DataTables.Item("DTLineas");

                //validar si ya se agregaron los lotes
                //result = ValidasiexisteloteentablaSalida(query);


                System.Console.WriteLine("Ejuctar query de formulario de salida de lotes");

                DT_LineasDocumento.ExecuteQuery(query);


                if (!DT_LineasDocumento.IsEmpty)
                {

                    //foreach (bool l in result)
                    //{
                    //    if (!l)
                    //    {
                    //        res = l;
                    //        break;
                    //    }
                    //}
                    //if (!res)
                    //{
                    System.Console.WriteLine("Formateo de campos de formulario de salida de lotes");
                    oForm.Items.Item("BtnAutSel").Enabled = false;
                    oForm.Items.Item("BtnMan").Enabled = false;
                    oForm.Items.Item("BtnAgre").Enabled = false;
                    oForm.Items.Item("BtnElim").Enabled = false;
                    grid_LineasDoc.SelectionMode = BoMatrixSelect.ms_None;
                    grid_LineasDoc.Item.Enabled = true;
                    grid_LineasDoc.Item.AffectsFormMode = false;


                    ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("ItemCode");
                    ColumnGrid.LinkedObjectType = "4";
                    ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("Almacen");
                    ColumnGrid.LinkedObjectType = "64";
                    if (Ini.Tipo.Equals("OWTR"))
                    {
                        ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("Almacen destino");
                        ColumnGrid.LinkedObjectType = "64";
                    }
                    ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("Sel");
                    ColumnGrid.Type = BoGridColumnType.gct_CheckBox;


                    for (int i = 0; i < grid_LineasDoc.Columns.Count; i++)
                    {

                        grid_LineasDoc.Columns.Item(i).Editable = false;
                        grid_LineasDoc.Columns.Item(i).TitleObject.Sortable = false;
                        grid_LineasDoc.Columns.Item(i).AffectsFormMode = false;
                    }

                    grid_LineasDoc.Columns.Item("Sel").Editable = true;

                    grid_LineasDoc.Columns.Item("ItemCode").TitleObject.Caption = "Número de articulo";
                    //grid_LineasDoc.Columns.Item("Cantidad solicitada").TitleObject.Caption = "Cantidad";
                    grid_LineasDoc.Columns.Item("Numero Doc").TitleObject.Caption = "Número documento";
                    grid_LineasDoc.Columns.Item("Tipo Doc").Visible = false;
                    grid_LineasDoc.AutoResizeColumns();
                    System.Console.WriteLine("Mostrar formulario de salida de lotes");
                    oForm.Visible = true;
                    //}
                    //else
                    //{
                    //    oForm.Close();
                    //}

                }
                else
                {
                    grid_LineasDoc.SelectionMode = BoMatrixSelect.ms_None;
                    grid_LineasDoc.Item.Enabled = false;
                    System.Console.WriteLine("Cerrar formulario de salida de lotes");
                    oForm.Close();
                    //res = true;

                }


                //return res;

            }
            catch (Exception ex)
            {
                oApp.SetStatusBarMessage("Error Add-On " + Ini.asm.GetName().Name + " : " + ex.Message.ToString() + " " + ex.StackTrace.ToString());
                //return false;

            }
            finally
            {
                if (oForm != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                }
                if (grid_LineasDoc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(grid_LineasDoc);
                }
                if (DT_LineasDocumento != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DT_LineasDocumento);
                }
                if (ColumnGrid != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ColumnGrid);
                }
            }
        }

        /// <summary>
        /// Funcion que agrega los lotes de la salida a la tabla de usuario desde boton autoseleccion
        /// </summary>
        /// <param name="pVal"></param>
        /// <returns></returns>
        internal void SalidaLotesAuto(ItemEvent pVal)
        {
            SAPbouiCOM.Form Salida = null;
            Grid grid_Disponible = null;
            Grid grid_LineasDoc = null;
            Grid grid_Sel = null;
            DataTable DT_Disponible = null;
            DataTable DT_LineasDocumento = null;
            DataTable DT_Sel = null;
            EditTextColumn ColumnGrid = null;
            SAPbobsCOM.Recordset oRs = null;

            string query = string.Empty;
            bool result = true, validacantidad = false, continuar = false;
            var Sb = new System.Text.StringBuilder();

            double cantidadSalida = 0, cantidaddesdeentrada = 0, cantidadRestante = 0, cantidadAcumulada = 0, val = 0;

            try
            {
                Salida = oApp.Forms.Item(pVal.FormUID);

                if (!Salida.Items.Item("BtnAutSel").Enabled)
                {
                    return;
                }

                Salida.Freeze(true);
                listaLotesSalida.Clear();
                grid_Disponible = (Grid)Salida.Items.Item("GrdDisp").Specific;
                grid_LineasDoc = (Grid)Salida.Items.Item("GrdLineas").Specific;
                grid_Sel = (Grid)Salida.Items.Item("GrdSel").Specific;
                DT_Disponible = (DataTable)Salida.DataSources.DataTables.Item("DTDispo");
                DT_LineasDocumento = (DataTable)Salida.DataSources.DataTables.Item("DTLineas");
                DT_Sel = (DataTable)Salida.DataSources.DataTables.Item("DTSel");
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


                for (int i = 0; i <= grid_LineasDoc.Rows.Count - 1; i++)
                {
                    Console.WriteLine(DT_LineasDocumento.GetValue("ItemCode", i));

                    //query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.InfoArticulosUDTSalidaValidacion.sql")).ReadToEnd();
                    //query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, DT_LineasDocumento.GetValue("ItemCode", i), DT_LineasDocumento.GetValue("Numero Doc", i), DT_LineasDocumento.GetValue("Tipo Doc", i)).ToString();
                    //oRs.DoQuery(query);
                    //if (oRs == null)
                    //{
                    //    result = false;
                    //    break;
                    //}
                    //else
                    //{
                    //    val = (int)oRs.Fields.Item(0).Value;
                    //    if (!val.Equals(0))
                    //    {
                    //        result = false;
                    //        break;
                    //    }
                    //}

                    Sb.Clear();
                    query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ValidaCantidadAcumuladadeloteSalida.sql")).ReadToEnd(); // trae la informacion de la cantidad total de lotes para el almacen, para controlar que exista la cantidad suficiente de lote en entrada
                    query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, DT_LineasDocumento.GetValue("ItemCode", i), DT_LineasDocumento.GetValue("Almacen", i), DT_LineasDocumento.GetValue("Cantidad", i)).ToString();

                    oRs.DoQuery(query);

                    if (oRs != null)
                    {
                        oRs.MoveFirst();
                        val = Convert.ToDouble(oRs.Fields.Item("Cantidad").Value);
                        if (val.Equals(0))
                        {
                            DT_Disponible.Clear();
                            Salida.Freeze(false);
                            oApp.MessageBox($"No hay lote cantidad suficiente para {DT_LineasDocumento.GetValue("ItemCode", i).ToString()} de los lotes ingresados en el almacen {DT_LineasDocumento.GetValue("Almacen", i).ToString()}", 1, "Ok");
                            result = false;
                            break;
                        }
                    }
                    else
                    {
                        DT_Disponible.Clear();
                        Salida.Freeze(false);
                        result = false;
                        break;
                    }



                    Sb.Clear();
                    query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.InfoArticulosconLoteUDT_Salida.sql")).ReadToEnd(); // trae la informacion de los lotes desde la tabla de entrada 
                    query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, DT_LineasDocumento.GetValue("ItemCode", i), DT_LineasDocumento.GetValue("Almacen", i)).ToString();

                    DT_Disponible.ExecuteQuery(query);

                    if (!DT_Disponible.IsEmpty)
                    {
                        cantidadSalida = Convert.ToDouble(DT_LineasDocumento.GetValue("Cantidad", i)); //cantidad que se va a sacar

                        cantidadRestante = 0;
                        cantidaddesdeentrada = 0;

                        for (int j = 0; j < DT_Disponible.Rows.Count; j++)
                        {
                            cantidaddesdeentrada = Convert.ToDouble(DT_Disponible.GetValue("Cantidad", j)); // cantidad que se ingreso con las entradas para ese item sin importar el lote

                            if (j.Equals(0))
                            {
                                if (cantidadSalida.Equals(cantidaddesdeentrada))
                                {
                                    cantidadRestante = cantidadSalida;
                                    cantidadAcumulada = 0;
                                    continuar = false;
                                }
                                else
                                {
                                    if (cantidadSalida < cantidaddesdeentrada)
                                    {
                                        cantidadRestante = cantidadSalida;
                                        cantidadAcumulada = 0;
                                        continuar = false;
                                    }
                                    else if (cantidadSalida > cantidaddesdeentrada)
                                    {
                                        cantidadRestante = cantidaddesdeentrada;
                                        cantidadAcumulada = cantidadSalida - cantidaddesdeentrada;
                                        continuar = true;
                                    }
                                }
                            }
                            else
                            {
                                if (cantidadAcumulada.Equals(cantidaddesdeentrada))
                                {
                                    cantidadRestante = cantidaddesdeentrada;
                                    cantidadAcumulada = 0;
                                    continuar = false;
                                }
                                else
                                {
                                    if (cantidadAcumulada < cantidaddesdeentrada)
                                    {
                                        cantidadRestante = cantidadAcumulada;
                                        cantidadAcumulada = 0;
                                        continuar = false;
                                    }
                                    else if (cantidadAcumulada > cantidaddesdeentrada)
                                    {
                                        cantidadRestante = cantidaddesdeentrada;
                                        cantidadAcumulada -= cantidaddesdeentrada;
                                        continuar = true;
                                    }
                                }
                            }
                            var lote = new LoteSalida { };
                            var fechavence = DT_Disponible.GetValue("Fecha vence", j);
                            DateTime fecha = Convert.ToDateTime(fechavence);

                            switch (Ini.Tipo)
                            {
                                case "OWTR":
                                    lote = new LoteSalida
                                    {
                                        ItemCode = DT_Disponible.GetValue("ItemCode", j).ToString(),
                                        NumeroLote = DT_Disponible.GetValue("Numero lote", j).ToString(),
                                        Almacen = DT_LineasDocumento.GetValue("Almacen", i).ToString(),
                                        AlmacenDestino = DT_LineasDocumento.GetValue("Almacen destino", i).ToString(),
                                        NumeroDocumento = DT_LineasDocumento.GetValue("Numero Doc", i).ToString(),
                                        TipoDocumento = DT_LineasDocumento.GetValue("Tipo Doc", i).ToString(),
                                        Cantidad = cantidadRestante,
                                        FechaProceso = fecha.ToString("dd/MM/yyyy"),
                                        NumLinea = (int)DT_LineasDocumento.GetValue("NumLinea", i),
                                        NumeroDocumentoEntrada = DT_Disponible.GetValue("Num Doc", j).ToString(),
                                        TipoDocumentoEntrada = DT_Disponible.GetValue("Tipo doc", j).ToString()


                                    };
                                    break;
                                default:
                                    lote = new LoteSalida
                                    {
                                        ItemCode = DT_Disponible.GetValue("ItemCode", j).ToString(),
                                        NumeroLote = DT_Disponible.GetValue("Numero lote", j).ToString(),
                                        Almacen = DT_LineasDocumento.GetValue("Almacen", i).ToString(),
                                        NumeroDocumento = DT_LineasDocumento.GetValue("Numero Doc", i).ToString(),
                                        TipoDocumento = DT_LineasDocumento.GetValue("Tipo Doc", i).ToString(),
                                        Cantidad = cantidadRestante,
                                        FechaProceso = fecha.ToString("dd/MM/yyyy"),
                                        NumLinea = (int)DT_LineasDocumento.GetValue("NumLinea", i),
                                        NumeroDocumentoEntrada = DT_Disponible.GetValue("Num Doc", j).ToString(),
                                        TipoDocumentoEntrada = DT_Disponible.GetValue("Tipo doc", j).ToString()

                                    };
                                    break;
                            }


                            listaLotesSalida.Add(lote);
                            SelectionManual = 2;
                            Salida.Items.Item("BtnMan").Enabled = false;
                            Salida.Items.Item("BtnElim").Enabled = false;
                            Sb.Clear();
                            if (!continuar) break;
                        }


                        //setDTDisponible(ref DT_Disponible, ref grid_Disponible, ref ColumnGrid);

                    }
                    else
                    {
                        Salida.Freeze(false);
                        oApp.MessageBox($"No hay lote disponible para {DT_LineasDocumento.GetValue("ItemCode", i).ToString()} de los lotes ingresados en el almacen {DT_LineasDocumento.GetValue("Almacen", i).ToString()}", 1, "Ok");
                        result = false;
                        break;
                    }

                    //grid_Disponible.Columns.Item("Fecha vence").TitleObject.Sortable = true;
                    //grid_Disponible.Columns.Item("Fecha vence").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);


                    //cantidadSalida = (int)DT_Disponible.GetValue("Cantidad", 0);
                    //cantidadEntrada = (int)DT_LineasDocumento.GetValue("Cantidad", i);

                    //if (cantidadEntrada < cantidadSalida)
                    //{
                    //    Salida.Freeze(false);
                    //    oApp.MessageBox($"No hay cantidad disponible para lote disponible para {DT_LineasDocumento.GetValue("ItemCode", i).ToString()} de los lotes ingresados", 1, "Ok");
                    //    result = false;
                    //    break;
                    //}





                    //Sb.Clear();

                    ////Agregar a la lista de update los valores de cada linea
                    //query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.UpdateInfoEntradasLotes.sql")).ReadToEnd();
                    //query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, (cantidadEntrada - cantidadSalida), DT_Disponible.GetValue("ItemCode", 0), DT_Disponible.GetValue("Tipo doc", 0), DT_Disponible.GetValue("Num Doc", 0), DT_Disponible.GetValue("Numero lote", 0)).ToString();

                    //ListaUpdateEntrada.Add(query);
                    //Sb.Clear();

                }

                //if (!result)
                //{
                //    DT_Disponible.Clear();
                //    DT_Sel.Clear();
                //    Salida.Freeze(false);
                //    return;
                //}
                if (cantidadAcumulada > 0)
                {
                    result = false;
                }

                DT_Disponible.Clear();
                Salida.Freeze(false);
                //Salida.Close();
                if (result)
                {
                    oApp.MessageBox($"Lotes agregados con exito", 1, "Ok");
                }
                grid_LineasDoc.Columns.Item("Sel").Editable = true;
            }
            catch (Exception ex)
            {
                Salida.Freeze(false);
                Console.WriteLine($"Error SalidaLotes {ex.Message} - {ex.StackTrace}");
            }
            finally
            {
                Salida.Freeze(false);
                Ini.ReleaseCOMObjetc(Salida);
                Ini.ReleaseCOMObjetc(grid_Disponible);
                Ini.ReleaseCOMObjetc(grid_LineasDoc);
                Ini.ReleaseCOMObjetc(grid_Sel);
                Ini.ReleaseCOMObjetc(DT_Disponible);
                Ini.ReleaseCOMObjetc(DT_LineasDocumento);
                Ini.ReleaseCOMObjetc(DT_Sel);
                Ini.ReleaseCOMObjetc(ColumnGrid);
                Ini.ReleaseCOMObjetc(oRs);
            }

        }


        private void setDTDisponible(ref DataTable dT_Disponible, ref Grid grid_Disponible, ref EditTextColumn columnGrid)
        {

            grid_Disponible.SelectionMode = BoMatrixSelect.ms_None;
            grid_Disponible.Item.Enabled = true;
            grid_Disponible.Item.AffectsFormMode = true;
            grid_Disponible.Columns.Item("ItemCode").Visible = false;
            grid_Disponible.Columns.Item("Tipo doc").Visible = false;
            grid_Disponible.Columns.Item("Num Doc").Visible = false;
            grid_Disponible.Columns.Item("Almacen").Visible = false;

            columnGrid = (EditTextColumn)grid_Disponible.Columns.Item("Sel");
            columnGrid.Type = BoGridColumnType.gct_CheckBox;


            for (int i = 0; i < grid_Disponible.Columns.Count; i++)
            {

                grid_Disponible.Columns.Item(i).AffectsFormMode = false;
                grid_Disponible.Columns.Item(i).TitleObject.Sortable = false;
                grid_Disponible.Columns.Item(i).Editable = false;
            }

            //grid_Disponible.Columns.Item("Cantidad").Editable = true;
            //grid_Disponible.Columns.Item("Sel").Editable = true;
            //grid_Disponible.Columns.Item("Sel").AffectsFormMode = true;
            grid_Disponible.Columns.Item("Fecha vence").Editable = false;
            grid_Disponible.Columns.Item("Fecha Admin").Editable = false;

            grid_Disponible.Columns.Item("Fecha vence").TitleObject.Sortable = true;
            grid_Disponible.Columns.Item("Fecha vence").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
            //grid_Disponible.Columns.Item("Cantidad").AffectsFormMode = true;


            grid_Disponible.AutoResizeColumns();
        }


        /// <summary>
        /// Funcion que actualiza la tabl CLC_LTS de entradas con las cantidades de las salidas
        /// </summary>
        /// <param name="listaUpdateEntrada"></param>
        private void UpdateTablaEntradaLotes(List<string> listaUpdateEntrada)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                foreach (var lista in listaUpdateEntrada)
                {
                    oRs.DoQuery(lista);
                    if (oRs == null)
                    {
                        Console.WriteLine("Error al ejecutar recordset de update entrada lotes");
                    }
                    else
                    {
                        Console.WriteLine("Exito al actualizar entrada lotes");
                    }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oRs);
            }
        }


        /// <summary>
        /// Funcion que agrega a la tabla CLS_LTSSA los datos de los lotes a sacar.
        /// </summary>
        /// <param name="listaLotesSalida"></param>
        private void AddLotesSalida(List<LoteSalida> listaLotesSalida)
        {
            SAPbobsCOM.UserTable Lotes = null;
            try
            {
                foreach (var l in listaLotesSalida)
                {
                    Console.WriteLine(
                        $"{l.ItemCode} | {l.Almacen} | {l.TipoDocumento} |{l.NumeroDocumento} | {l.NumeroLote} | {l.Cantidad} | {l.FechaProceso}"
                    );

                    Lotes = oCompany.UserTables.Item("CLC_LTSSA");

                    Lotes.Name = l.ItemCode;
                    Lotes.UserFields.Fields.Item("U_TipDoc").Value = l.TipoDocumento;
                    Lotes.UserFields.Fields.Item("U_NumDoc").Value = l.NumeroDocumento;
                    Lotes.UserFields.Fields.Item("U_NumLot").Value = l.NumeroLote;
                    Lotes.UserFields.Fields.Item("U_Cantidad").Value = l.Cantidad;
                    Lotes.UserFields.Fields.Item("U_FchSal").Value = l.FechaProceso;
                    Lotes.UserFields.Fields.Item("U_Almacen").Value = l.Almacen;

                    if (Lotes.Add().Equals(0))
                    {
                        Console.WriteLine("Datos agregados a tabla lotes salida exitosamente");

                    }
                    else
                    {
                        Console.WriteLine($"Error al agregar datos a tabla de lotes salida. {oCompany.GetLastErrorDescription()} ");

                    }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(Lotes);
            }

        }


        /// <summary>
        /// Funcion para mostrar informacion de lotes para salida cuando se selecciona manualmente el check de lotes de las lienas del documento
        /// </summary>
        /// <param name="pVal"></param>
        internal void AddInfoLotesSalida(ItemEvent pVal)
        {
            SAPbouiCOM.Form Salida = null;
            Grid grid_Disponible = null;
            Grid grid_LineasDoc = null;
            Grid grid_Sel = null;
            DataTable DT_Disponible = null;
            DataTable DT_LineasDocumento = null;
            DataTable DT_Sel = null;
            //EditTextColumn ColumnGrid = null;

            string query = string.Empty, Tipo = string.Empty;
            var Sb = new System.Text.StringBuilder();



            try
            {
                Salida = oApp.Forms.Item(pVal.FormUID);
                Salida.Freeze(true);

                grid_Disponible = (Grid)Salida.Items.Item("GrdDisp").Specific;
                grid_LineasDoc = (Grid)Salida.Items.Item("GrdLineas").Specific;
                grid_Sel = (Grid)Salida.Items.Item("GrdSel").Specific;
                DT_Disponible = (DataTable)Salida.DataSources.DataTables.Item("DTDispo");
                DT_LineasDocumento = (DataTable)Salida.DataSources.DataTables.Item("DTLineas");
                DT_Sel = (DataTable)Salida.DataSources.DataTables.Item("DTSel");

                //if (DT_Disponible.GetValue("Sel", pVal.Row).Equals("N"))
                //{
                //    DT_Sel.Clear();
                //    Salida.Freeze(false);
                //    return;
                //}

                //Console.WriteLine(DT_Disponible.GetValue("ItemCode", pVal.Row));

                for (int i = 0; i <= grid_LineasDoc.Rows.Count - 1; i++)
                {
                    if (DT_LineasDocumento.GetValue("Sel", i).ToString().Equals("Y"))
                    {
                        Tipo = DT_LineasDocumento.GetValue("Tipo Doc", i).ToString();
                    }
                }


                query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.InfoArticulosUDTSalida.sql")).ReadToEnd();
                query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, DT_Disponible.GetValue("ItemCode", pVal.Row), DT_Disponible.GetValue("Numero lote", pVal.Row), Tipo).ToString();


                DT_Sel.ExecuteQuery(query);

                //if (DT_Sel.IsEmpty)
                //{
                //    Sb.Clear();
                //    query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.InfoArticulosLoteDisponible.sql")).ReadToEnd();
                //    query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, DT_Disponible.GetValue("ItemCode", pVal.Row), DT_Disponible.GetValue("Tipo doc", pVal.Row), DT_Disponible.GetValue("Num Doc", pVal.Row), DT_Disponible.GetValue("Numero lote", pVal.Row), DT_Disponible.GetValue("Cantidad", pVal.Row)).ToString();
                //    DT_Sel.ExecuteQuery(query);
                //}

                if (!DT_Sel.IsEmpty)
                {
                    grid_Sel.SelectionMode = BoMatrixSelect.ms_None;
                    grid_Sel.Item.Enabled = false;
                    grid_Sel.Item.AffectsFormMode = true;
                    grid_Sel.Columns.Item("ItemCode").Visible = false;
                    grid_Sel.Columns.Item("Tipo doc").Visible = false;
                    grid_Sel.Columns.Item("Num Doc").Visible = false;



                    for (int i = 0; i < grid_Sel.Columns.Count; i++)
                    {

                        grid_Sel.Columns.Item(i).AffectsFormMode = false;
                        grid_Sel.Columns.Item(i).TitleObject.Sortable = false;
                        grid_Sel.Columns.Item(i).Editable = false;
                    }
                    grid_Sel.AutoResizeColumns();
                }
                else
                {
                    grid_Sel.SelectionMode = BoMatrixSelect.ms_None;
                    grid_Sel.Item.Enabled = false;
                }

                Salida.Freeze(false);
            }
            catch (Exception ex)
            {
                Salida.Freeze(false);
                Console.WriteLine($"Error SalidaLotes {ex.Message} - {ex.StackTrace}");
            }
            finally
            {
                Ini.ReleaseCOMObjetc(Salida);
                Ini.ReleaseCOMObjetc(grid_Disponible);
                Ini.ReleaseCOMObjetc(grid_LineasDoc);
                Ini.ReleaseCOMObjetc(grid_Sel);
                Ini.ReleaseCOMObjetc(DT_Disponible);
                Ini.ReleaseCOMObjetc(DT_LineasDocumento);
                Ini.ReleaseCOMObjetc(DT_Sel);
            }
        }

        /// <summary>
        /// Habilitar la opcion de seleccion manual 
        /// </summary>
        /// <param name="pVal">Variable con valores de evento</param>
        internal void EnableSelectionManual(ItemEvent pVal)
        {

            SAPbouiCOM.Form Salida = null;
            Grid grid_Disponible = null;
            Grid grid_LineasDoc = null;
            DataTable DT_Disponible = null;

            string query = string.Empty;
            var Sb = new System.Text.StringBuilder();

            try
            {
                Salida = oApp.Forms.Item(pVal.FormUID);
                Salida.Freeze(true);
                if (Salida.Items.Item("BtnMan").Enabled)
                {
                    grid_Disponible = (Grid)Salida.Items.Item("GrdDisp").Specific;
                    DT_Disponible = (DataTable)Salida.DataSources.DataTables.Item("DTDispo");
                    grid_LineasDoc = (Grid)Salida.Items.Item("GrdLineas").Specific;

                    grid_Disponible.Item.Enabled = true;
                    grid_LineasDoc.Columns.Item("Sel").Editable = true;
                    Salida.Items.Item("BtnAutSel").Enabled = false;

                    SelectionManual = 1;

                    query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.SinDatos.sql")).ReadToEnd();
                    query = Sb.AppendFormat(CultureInfo.InvariantCulture, query).ToString();

                    DT_Disponible.ExecuteQuery(query);
                    grid_Disponible.AutoResizeColumns();
                    grid_Disponible.Columns.Item("RowsHeader").Width = 20;
                }


            }
            catch (Exception ex)
            {
                Salida.Freeze(false);
                oApp.SetStatusBarMessage("Error Add-On " + Ini.asm.GetName().Name + " : " + ex.Message.ToString() + " " + ex.StackTrace.ToString());
            }
            finally
            {
                Salida.Freeze(false);

                Ini.ReleaseCOMObjetc(Salida);
                Ini.ReleaseCOMObjetc(grid_Disponible);
                Ini.ReleaseCOMObjetc(grid_LineasDoc);
                Ini.ReleaseCOMObjetc(DT_Disponible);

            }
        }

        /// <summary>
        /// Adicionar manual los lotes en el grid de seleccionados
        /// </summary>
        /// <param name="pVal"></param>
        internal void AddLoteSel(ItemEvent pVal)
        {
            SAPbouiCOM.Form Salida = null;
            Grid grid_Disponible = null;
            Grid grid_LineasDoc = null;
            DataTable DT_Disponible = null;
            Grid grid_Seleccion = null;
            DataTable DT_Seleccion = null;

            string query = string.Empty;
            var Sb = new System.Text.StringBuilder();
            List<string> ListaLotesAdd = new List<string>();

            int lineaseleccionada = 0;
            double cantidadDispo = 0, cantidadSel = 0;

            try
            {
                Salida = oApp.Forms.Item(pVal.FormUID);
                Salida.Freeze(true);

                if (!Salida.Items.Item("BtnMan").Enabled)
                {
                    return;
                }
                grid_Disponible = (Grid)Salida.Items.Item("GrdDisp").Specific;
                DT_Disponible = (DataTable)Salida.DataSources.DataTables.Item("DTDispo");
                grid_LineasDoc = (Grid)Salida.Items.Item("GrdLineas").Specific;
                grid_Seleccion = (Grid)Salida.Items.Item("GrdSel").Specific;
                DT_Seleccion = (DataTable)Salida.DataSources.DataTables.Item("DTSel");

                lineaseleccionada = listaLotesSalida.Count;


                for (int i = 0; i < DT_Disponible.Rows.Count; i++)
                {
                    if (DT_Disponible.GetValue("Sel", i).Equals("Y"))
                    {
                        //Validar que la cantidad ingresada no sea mayor a la cantidad disponible
                        var catidaDisponible = DT_Disponible.GetValue("Cantidad disponible", i);
                        var catidaIngresada = DT_Disponible.GetValue("Cantidad", i);

                        cantidadDispo = Convert.ToDouble(catidaDisponible);
                        cantidadSel = Convert.ToDouble(catidaIngresada);

                        if (cantidadDispo < cantidadSel)
                        {
                            oApp.MessageBox($"La cantidad ingresada es mayor a la cantidad disponible.", 1, "Ok");
                            return;
                        }

                        if (cantidadSel > 0)
                        {
                            Sb.Clear();
                            query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.InfoLotesAddmanual.sql")).ReadToEnd();
                            var fechaVence = DT_Disponible.GetValue("Fecha vence", i);
                            DateTime ofechaVence = Convert.ToDateTime(fechaVence);

                            var fechaAdmin = DT_Disponible.GetValue("Fecha Admin", i);
                            DateTime ofechaAdim = Convert.ToDateTime(fechaVence);

                            query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, DT_Disponible.GetValue("ItemCode", i), DT_Disponible.GetValue("Tipo doc", i), DT_Disponible.GetValue("Num Doc", i), DT_Disponible.GetValue("Numero lote", i), DT_Disponible.GetValue("Cantidad", i), ofechaVence.ToString("dd/MM/yyyy"), ofechaAdim.ToString("dd/MM/yyyy"), DT_Disponible.GetValue("Almacen", i), lineaseleccionada, i).ToString();
                            ListaLotesAdd.Add(query);
                            lineaseleccionada++;

                        }
                    }
                }


                if (ListaLotesAdd.Any())
                {
                    query = string.Join(" UNION ALL ", ListaLotesAdd);
                    DT_Seleccion.ExecuteQuery(query);

                    if (!DT_Seleccion.IsEmpty)
                    {
                        grid_Seleccion.SelectionMode = BoMatrixSelect.ms_None;
                        grid_Seleccion.Item.Enabled = false;
                        grid_Seleccion.Item.AffectsFormMode = true;
                        grid_Seleccion.Columns.Item("ItemCode").Visible = false;
                        grid_Seleccion.Columns.Item("Tipo doc").Visible = false;
                        grid_Seleccion.Columns.Item("Num Doc").Visible = false;
                        grid_Seleccion.Columns.Item("NumLinea").Visible = false;



                        for (int i = 0; i < grid_Seleccion.Columns.Count; i++)
                        {

                            grid_Seleccion.Columns.Item(i).AffectsFormMode = false;
                            grid_Seleccion.Columns.Item(i).TitleObject.Sortable = false;
                            grid_Seleccion.Columns.Item(i).Editable = false;
                        }
                        grid_Seleccion.SelectionMode = BoMatrixSelect.ms_Single;
                        grid_Seleccion.AutoResizeColumns();
                    }
                    else
                    {
                        grid_Seleccion.SelectionMode = BoMatrixSelect.ms_None;
                        grid_Seleccion.Item.Enabled = false;
                    }
                }
            }
            catch (Exception ex)
            {
                Salida.Freeze(false);
                oApp.SetStatusBarMessage("Error Add-On " + Ini.asm.GetName().Name + " : " + ex.Message.ToString() + " " + ex.StackTrace.ToString());
            }
            finally
            {
                Salida.Freeze(false);

                Ini.ReleaseCOMObjetc(Salida);
                Ini.ReleaseCOMObjetc(grid_Disponible);
                Ini.ReleaseCOMObjetc(grid_LineasDoc);
                Ini.ReleaseCOMObjetc(DT_Disponible);
                Ini.ReleaseCOMObjetc(grid_Seleccion);
                Ini.ReleaseCOMObjetc(DT_Seleccion);

            }
        }

        /// <summary>
        /// Metodo para eliminar fila de lote seleccionado
        /// </summary>
        /// <param name="pVal">Variable con valores del evento</param>
        internal void DelLoteSel(ItemEvent pVal)
        {
            SAPbouiCOM.Form Salida = null;
            Grid grid_Disponible = null;
            DataTable DT_Disponible = null;
            Grid grid_Seleccion = null;
            DataTable DT_Seleccion = null;

            try
            {
                Salida = oApp.Forms.Item(pVal.FormUID);
                Salida.Freeze(true);

                if (!Salida.Items.Item("BtnMan").Enabled)
                {
                    return;
                }

                grid_Disponible = (Grid)Salida.Items.Item("GrdDisp").Specific;
                DT_Disponible = (DataTable)Salida.DataSources.DataTables.Item("DTDispo");
                grid_Seleccion = (Grid)Salida.Items.Item("GrdSel").Specific;
                DT_Seleccion = (DataTable)Salida.DataSources.DataTables.Item("DTSel");

                for (int i = 0; i < grid_Seleccion.Rows.Count; i++)
                {
                    if (grid_Seleccion.Rows.IsSelected(i))
                    {

                        if (listaLotesSalida.Any())
                        {
                            listaLotesSalida.RemoveAll(x =>
                                                    x.ItemCode == DT_Seleccion.GetValue("ItemCode", i).ToString() &&
                                                    x.NumeroLote == DT_Seleccion.GetValue("Numero lote", i).ToString() &&
                                                    x.Almacen == DT_Seleccion.GetValue("Almacen", i).ToString() &&
                                                    x.NumLinea == (int)DT_Seleccion.GetValue("NumLinea", i));

                            //listaLotesSalida.RemoveAt((int)DT_Seleccion.GetValue("Index", i));

                        }
                        DT_Seleccion.Rows.Remove(i);                        

                        if (!grid_Disponible.Item.Enabled)
                        {
                            grid_Disponible.Item.Enabled = true;

                        }

                        Salida.Mode = BoFormMode.fm_UPDATE_MODE;

                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Salida.Freeze(false);
                oApp.SetStatusBarMessage("Error Add-On " + Ini.asm.GetName().Name + " : " + ex.Message.ToString() + " " + ex.StackTrace.ToString());

            }
            finally
            {
                Salida.Freeze(false);

                Ini.ReleaseCOMObjetc(Salida);
                Ini.ReleaseCOMObjetc(grid_Disponible);
                Ini.ReleaseCOMObjetc(DT_Disponible);
                Ini.ReleaseCOMObjetc(grid_Seleccion);
                Ini.ReleaseCOMObjetc(DT_Seleccion);
            }

            
        }

        /// <summary>
        /// Metodo para validar cantidades y lotes seleccionados pra luego agregarlos a la lista de lotes salida
        /// </summary>
        /// <param name="pVal">Variable con valores de evento</param>
        /// <returns></returns>
        internal bool AddLoteSelToList(ItemEvent pVal)
        {
            SAPbouiCOM.Form Salida = null;
            Grid grid_LineasDoc = null;
            DataTable DT_Lineas = null;
            Grid grid_Seleccion = null;
            DataTable DT_Seleccion = null;
            Grid grid_Disponible = null;
            DataTable DT_Disponible = null;

            double cantidadLineas = 0, cantidadDisponible = 0, cantidadDisponibleSelecionada = 0, cantidadDisponibleLinea = 0, cantidadDisponibleSelecionadaLinea = 0;

            bool result = true;

            string LoteDisponibleLinea = string.Empty, LoteSeleccionadoLinea = string.Empty;

            try
            {
                Salida = oApp.Forms.Item(pVal.FormUID);

                if (Salida.Mode.Equals(BoFormMode.fm_UPDATE_MODE))
                {
                    if (SelectionManual.Equals(1))
                    {
                        DT_Lineas = (DataTable)Salida.DataSources.DataTables.Item("DTLineas");
                        grid_LineasDoc = (Grid)Salida.Items.Item("GrdLineas").Specific;
                        grid_Seleccion = (Grid)Salida.Items.Item("GrdSel").Specific;
                        DT_Seleccion = (DataTable)Salida.DataSources.DataTables.Item("DTSel");
                        grid_Disponible = (Grid)Salida.Items.Item("GrdDisp").Specific;
                        DT_Disponible = (DataTable)Salida.DataSources.DataTables.Item("DTDispo");

                        //Capturar el valor principal de cantidad de la salida
                        for (int i = 0; i <= grid_LineasDoc.Rows.Count - 1; i++)
                        {
                            if (DT_Lineas.GetValue("Sel", i).ToString().Equals("Y"))
                            {
                                cantidadLineas = Convert.ToDouble(DT_Lineas.GetValue("Cantidad", i));
                                break;
                            }

                        }

                        //Capturar el valor de cantidad de lote ingresado como disponible en seleccionados
                        for (int i = 0; i <= grid_Seleccion.Rows.Count - 1; i++)
                        {

                            cantidadDisponibleSelecionada += Convert.ToDouble(DT_Seleccion.GetValue("Cantidad", i));

                        }

                        //Capturar el valor de cantidad de lote ingresado como disponible en disponibles
                        for (int i = 0; i <= grid_Disponible.Rows.Count - 1; i++)
                        {
                            if (DT_Disponible.GetValue("Sel", i).ToString().Equals("Y"))
                            {
                                cantidadDisponible += Convert.ToDouble(DT_Disponible.GetValue("Cantidad", i));
                            }
                        }

                        //Validar que las cantidades entre principal y disponble ingresada sean iguales
                        if (!cantidadLineas.Equals(cantidadDisponible))
                        {
                            oApp.MessageBox("La cantidad del lote ingresado debe ser igual a la cantidad del documento.", 1, "Ok");
                            result = false;
                            return result;
                        }

                        //Validar que las cantidades entre Disponible y Selecionadas ingresada sean iguales
                        if (!cantidadDisponibleSelecionada.Equals(cantidadDisponible))
                        {
                            oApp.MessageBox("La cantidad del lote ingresado en disponibles debe ser igual a la cantidad del lote ingresado en seleccionados.", 1, "Ok");
                            result = false;
                            return result;
                        }


                        //validar que las cantidades entre disponible y seleccionados sean iguales por linea de cada lote
                        for (int i = 0; i <= grid_Disponible.Rows.Count - 1; i++)
                        {
                            if (DT_Disponible.GetValue("Sel", i).ToString().Equals("Y"))
                            {
                                cantidadDisponibleLinea = Convert.ToDouble(DT_Disponible.GetValue("Cantidad", i));
                                LoteDisponibleLinea = DT_Disponible.GetValue("Numero lote", i).ToString();

                                for (int j = 0; j <= grid_Seleccion.Rows.Count - 1; j++)
                                {
                                    var numlinea = (int)DT_Seleccion.GetValue("NumLinea", j);
                                    if (numlinea.Equals(i))
                                    {
                                        cantidadDisponibleSelecionadaLinea = Convert.ToDouble(DT_Seleccion.GetValue("Cantidad", j));
                                        LoteSeleccionadoLinea = DT_Seleccion.GetValue("Numero lote", j).ToString();
                                        break;
                                    }

                                }
                                if (!cantidadDisponibleLinea.Equals(cantidadDisponibleSelecionadaLinea))
                                {
                                    oApp.MessageBox("La cantidad del lote ingresado en disponibles por líneas difiere de lo ingresado en seleccionados.", 1, "Ok");
                                    result = false;
                                    return result;
                                }

                                //Validar que los lotes sean iguales por linea Disponible y seleccionada
                                if (!LoteDisponibleLinea.Equals(LoteSeleccionadoLinea))
                                {
                                    oApp.MessageBox("Existe diferencia entre los lotes seleccionados en disponible y los agregados en seleccionados.", 1, "Ok");
                                    result = false;
                                    return result;
                                }
                            }
                        }

                        ////validar que los lotes entre disponible y seleccionados sean iguales por linea
                        //for (int i = 0; i <= grid_Disponible.Rows.Count - 1; i++)
                        //{
                        //    if (DT_Disponible.GetValue("Sel", i).ToString().Equals("Y"))
                        //    {
                        //        LoteDisponibleLinea = DT_Disponible.GetValue("Numero lote", i).ToString();

                        //        existelote = false;
                        //        for (int j = 0; j <= grid_Seleccion.Rows.Count - 1; j++)
                        //        {

                        //            LoteSeleccionadoLinea = DT_Seleccion.GetValue("Numero lote", j).ToString();
                        //            //Validar que los lotes sean iguales por linea Disponible y seleccionada

                        //            if (LoteDisponibleLinea.Equals(LoteSeleccionadoLinea))
                        //            {
                        //                existelote = true;
                        //                break;
                        //            }                                    
                        //        }
                        //    }
                        //}

                        //if (!existelote)
                        //{
                        //    oApp.MessageBox("Existe diferencia entre los lotes seleccionados en disponible y los agregados en seleccionados.", 1, "Ok");
                        //    result = false;
                        //    return result;
                        //}


                        AddLotesSalidaManualToLits(DT_Seleccion, grid_LineasDoc, DT_Lineas);
                    }
                }


                return result;
            }
            catch (Exception ex)
            {

                oApp.MessageBox($"Error: Add-On {Ini.asm.GetName().Name} - {ex.Message}", 1, "Ok");
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(Salida);
                Ini.ReleaseCOMObjetc(grid_LineasDoc);
                Ini.ReleaseCOMObjetc(DT_Lineas);
                Ini.ReleaseCOMObjetc(grid_Seleccion);
                Ini.ReleaseCOMObjetc(DT_Seleccion);
                Ini.ReleaseCOMObjetc(grid_Disponible);
                Ini.ReleaseCOMObjetc(DT_Disponible);
            }
        }

        /// <summary>
        /// Metodo para agregar los lotes selecionados a la lista de salida
        /// </summary>
        /// <param name="dT_Seleccion">Data table SAP con información de lotes de salida</param>
        private void AddLotesSalidaManualToLits(DataTable dT_Seleccion, Grid grid_LineasDoc, DataTable dT_Lineas)
        {
            var tipo = Ini.Tipo;


            for (int i = 0; i <= grid_LineasDoc.Rows.Count - 1; i++)
            {
                if (dT_Lineas.GetValue("Sel", i).ToString().Equals("Y"))
                {
                    for (int j = 0; j < dT_Seleccion.Rows.Count; j++)
                    {

                        var lote = new LoteSalida { };
                        var fechavence = dT_Seleccion.GetValue("Fecha vence", j);
                        DateTime fecha = Convert.ToDateTime(fechavence);

                        switch (Ini.Tipo)
                        {
                            case "OWTR":
                                lote = new LoteSalida
                                {
                                    ItemCode = dT_Seleccion.GetValue("ItemCode", j).ToString(),
                                    NumeroLote = dT_Seleccion.GetValue("Numero lote", j).ToString(),
                                    Almacen = dT_Lineas.GetValue("Almacen", i).ToString(),
                                    AlmacenDestino = dT_Lineas.GetValue("Almacen destino", i).ToString(),
                                    NumeroDocumento = dT_Lineas.GetValue("Numero Doc", i).ToString(),
                                    TipoDocumento = dT_Lineas.GetValue("Tipo Doc", i).ToString(),
                                    Cantidad = Convert.ToDouble(dT_Seleccion.GetValue("Cantidad", j).ToString()),
                                    FechaProceso = fecha.ToString("dd/MM/yyyy"),
                                    NumLinea = (int)dT_Lineas.GetValue("NumLinea", i),
                                    NumeroDocumentoEntrada = dT_Seleccion.GetValue("Num Doc", j).ToString(),
                                    TipoDocumentoEntrada = dT_Seleccion.GetValue("Tipo doc", j).ToString()


                                };
                                break;
                            default:
                                lote = new LoteSalida
                                {
                                    ItemCode = dT_Seleccion.GetValue("ItemCode", j).ToString(),
                                    NumeroLote = dT_Seleccion.GetValue("Numero lote", j).ToString(),
                                    Almacen = dT_Lineas.GetValue("Almacen", i).ToString(),
                                    NumeroDocumento = dT_Lineas.GetValue("Numero Doc", i).ToString(),
                                    TipoDocumento = dT_Lineas.GetValue("Tipo Doc", i).ToString(),
                                    Cantidad = Convert.ToDouble(dT_Seleccion.GetValue("Cantidad", j).ToString()),
                                    FechaProceso = fecha.ToString("dd/MM/yyyy"),
                                    NumLinea = (int)dT_Lineas.GetValue("NumLinea", i),
                                    NumeroDocumentoEntrada = dT_Seleccion.GetValue("Num Doc", j).ToString(),
                                    TipoDocumentoEntrada = dT_Seleccion.GetValue("Tipo doc", j).ToString()

                                };
                                break;
                        }


                        listaLotesSalida.Add(lote);

                    }
                }
            }
        }

        internal void RollBack(ItemEvent pVal, string Tipo)
        {
            SAPbouiCOM.Form Salida = null;
            SAPbouiCOM.Matrix LineasSalida = null;
            EditText ItemCode = null;
            EditText oDocNum = null;
            SAPbobsCOM.Recordset oRs = null, oRsUpdate = null;

            string query = string.Empty;
            var Sb = new System.Text.StringBuilder();

            try
            {
                Console.WriteLine("Inicio de rollback de articulos con lote entrada");

                Salida = oApp.Forms.Item(pVal.FormUID);
                LineasSalida = (SAPbouiCOM.Matrix)Salida.Items.Item("38").Specific;
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRsUpdate = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //dBDataSourceLineasDocumento = Entrada.DataSources.DBDataSources.Item("PDN1");


                for (int i = 1; i <= LineasSalida.RowCount; i++)
                {
                    ItemCode = (SAPbouiCOM.EditText)LineasSalida.Columns.Item("1").Cells.Item(i).Specific;

                    if (!string.IsNullOrEmpty(ItemCode.String))
                    {
                        oDocNum = (SAPbouiCOM.EditText)Salida.Items.Item("8").Specific;

                        query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.SelectLoteSalidasRollback.sql")).ReadToEnd();
                        query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, ItemCode.String, oDocNum.String, Tipo).ToString();
                        oRs.DoQuery(query);
                        System.Console.WriteLine($"Query select lote rollback salida {query}");
                        if (oRs == null)
                        {
                            System.Console.WriteLine($"Error con recordset query select lote rollback salida {query}");
                        }
                        else
                        {
                            oRs.MoveFirst();
                            while (!oRs.EoF)
                            {
                                Sb.Clear();
                                query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.UpdateEntradaSalidasRollback.sql")).ReadToEnd();
                                query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, oRs.Fields.Item("U_NumLot"), oRs.Fields.Item("U_Cantidad")).ToString();
                                oRsUpdate.DoQuery(query);
                                if (oRsUpdate == null)
                                {
                                    System.Console.WriteLine($"Error con recordset update query rollback salida, {query}");
                                }
                                else
                                {
                                    System.Console.WriteLine($"Update de entrada exitoso rollback Salida, {query}");

                                }
                            }
                            Sb.Clear();
                        }

                        query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.DeleteArticulosconLoteSalida.sql")).ReadToEnd();
                        query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, ItemCode.String, Tipo, oDocNum.String).ToString();
                        oRs.DoQuery(query);

                        Sb.Clear();
                        System.Console.WriteLine($"Query rollback Salida delete, {query}");


                    }
                }

                Console.WriteLine("Fin de rollback de articulos con lote Salida");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en rollback de articulos con lote Salida, {ex.Message}");
            }
            finally
            {
                Ini.ReleaseCOMObjetc(Salida);
                Ini.ReleaseCOMObjetc(LineasSalida);
                Ini.ReleaseCOMObjetc(ItemCode);
                Ini.ReleaseCOMObjetc(oDocNum);
                Ini.ReleaseCOMObjetc(oRs);
                Ini.ReleaseCOMObjetc(oRsUpdate);
            }
        }

        internal void formClose()
        {
            try
            {
                listaLotesSalida.Clear();
                SelectionManual = 0;

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en formClose lotes de articulos con lote salida, {ex.Message}");
            }
            finally
            {

            }
        }

        /// <summary>
        /// Inserta en CLC_LTSSA las filas de listaLotesSalida.
        /// </summary>
        /// <param name="estado">
        /// "P" = Pendiente (borrador en aprobación): SOLO se guarda el registro,
        ///       NO se descuenta cantidad del lote de origen todavía (el documento
        ///       puede ser rechazado, no hay que tocar el saldo disponible).
        /// "C" = Confirmado (documento definitivo, flujo normal sin aprobación):
        ///       se guarda Y se descuenta la cantidad del lote de origen, igual
        ///       que siempre se hizo.
        /// </param>
        internal bool SalidaLotes(string estado = "C")
        {
            bool res = true;
            SAPbobsCOM.UserTable Lotes = null;
            SAPbobsCOM.UserTable LotesEntrada = null;
            List<string> ListaUpdateEntrada = new List<string>();
            SAPbobsCOM.Recordset oRs = null;
            var Sb = new System.Text.StringBuilder();
            var query = string.Empty;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                for (int i = 0; i < listaLotesSalida.Count; i++)
                {
                    var item = listaLotesSalida[i];

                    Lotes = oCompany.UserTables.Item("CLC_LTSSA");

                    Lotes.Name = item.ItemCode;
                    Lotes.UserFields.Fields.Item("U_TipDoc").Value = item.TipoDocumento;
                    Lotes.UserFields.Fields.Item("U_NumDoc").Value = item.NumeroDocumento;
                    Lotes.UserFields.Fields.Item("U_NumLot").Value = item.NumeroLote;
                    Lotes.UserFields.Fields.Item("U_Cantidad").Value = item.Cantidad;
                    Lotes.UserFields.Fields.Item("U_FchSal").Value = DateTime.Now;
                    Lotes.UserFields.Fields.Item("U_almacen").Value = item.Almacen;
                    Lotes.UserFields.Fields.Item("U_Estado").Value = estado;
                    Lotes.UserFields.Fields.Item("U_TDocEnt").Value = item.TipoDocumentoEntrada ?? string.Empty;
                    Lotes.UserFields.Fields.Item("U_NDocEnt").Value = item.NumeroDocumentoEntrada ?? string.Empty;
                    Lotes.UserFields.Fields.Item("U_AlmDest").Value = item.AlmacenDestino ?? string.Empty;
                    Lotes.UserFields.Fields.Item("U_FchProc").Value = item.FechaProceso ?? string.Empty;

                    if (Lotes.Add().Equals(0))
                    {
                        Console.WriteLine($"Datos agregados a tabla lotes exitosamente ({estado}) {item.ItemCode} - {item.TipoDocumento} - {item.NumeroDocumento} - {item.NumeroLote}");

                        if (estado.Equals("P"))
                        {
                            // Pendiente de aprobación: no se descuenta cantidad ni se
                            // crea la entrada de traslado todavía. Eso se hace en
                            // ConfirmarPendientesSalida() cuando el documento se
                            // confirme como definitivo.
                            continue;
                        }

                        if (Ini.Tipo.Equals("OWTR"))
                        {
                            Sb.Clear();

                            query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.AlmacenManejaLote.sql")).ReadToEnd();
                            query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, item.AlmacenDestino).ToString();
                            oRs.DoQuery(query);
                            if (oRs != null)
                            {
                                oRs.MoveFirst();

                                string almacenDes = oRs.Fields.Item(0).Value.ToString();
                                if (!String.IsNullOrEmpty(almacenDes))
                                {
                                    // CORRECCION: si el traslado tiene 2 líneas con el mismo
                                    // artículo+lote hacia el mismo almacén destino, se SUMA la
                                    // cantidad a la fila que ya existe en vez de duplicarla.
                                    string queryBuscaDestino = $"SELECT \"Code\" FROM \"@CLC_LTS\" WHERE \"Name\" = '{item.ItemCode}' " +
                                                                $"AND \"U_TipDoc\" = '{item.TipoDocumento}' AND \"U_NumDoc\" = {item.NumeroDocumento} AND \"U_NumLot\" = '{item.NumeroLote}' " +
                                                                $"AND \"U_almacen\" = '{item.AlmacenDestino}'";
                                    oRs.DoQuery(queryBuscaDestino);

                                    if (oRs.RecordCount > 0)
                                    {
                                        oRs.MoveFirst();
                                        string codeExistenteDestino = oRs.Fields.Item("Code").Value.ToString();
                                        oRs.DoQuery($"UPDATE \"@CLC_LTS\" SET \"U_CantLote\" = (\"U_CantLote\" + {item.Cantidad}) WHERE \"Code\" = '{codeExistenteDestino}'");
                                        Console.WriteLine($"Lote destino ya existía en el traslado, se sumó cantidad {item.ItemCode} - {item.TipoDocumento} - {item.NumeroDocumento} - {item.NumeroLote}");
                                    }
                                    else
                                    {
                                        LotesEntrada = oCompany.UserTables.Item("CLC_LTS");
                                        LotesEntrada.Name = item.ItemCode;
                                        LotesEntrada.UserFields.Fields.Item("U_TipDoc").Value = item.TipoDocumento;
                                        LotesEntrada.UserFields.Fields.Item("U_NumDoc").Value = item.NumeroDocumento;
                                        LotesEntrada.UserFields.Fields.Item("U_NumLot").Value = item.NumeroLote;
                                        LotesEntrada.UserFields.Fields.Item("U_CantLote").Value = item.Cantidad;
                                        LotesEntrada.UserFields.Fields.Item("U_FchVenLt").Value = item.FechaProceso;
                                        LotesEntrada.UserFields.Fields.Item("U_FchAdmLt").Value = DateTime.Now;
                                        LotesEntrada.UserFields.Fields.Item("U_almacen").Value = item.AlmacenDestino;
                                        if (LotesEntrada.Add().Equals(0))
                                        {
                                            Console.WriteLine($"Datos agregados a tabla lotes entrada exitosamente con traslado {item.ItemCode} - {item.TipoDocumento} - {item.NumeroDocumento} - {item.NumeroLote}");
                                        }
                                        else
                                        {
                                            Console.WriteLine($"Error al agregar datos a tabla de lotes entrada con traslado. {oCompany.GetLastErrorDescription()}. {item.ItemCode} - {item.TipoDocumento} - {item.NumeroDocumento} - {item.NumeroLote}");
                                            res = false;
                                            break;
                                        }
                                    }
                                }

                            }

                        }
                    }
                    else
                    {
                        Console.WriteLine($"Error al agregar datos a tabla de lotes. {oCompany.GetLastErrorDescription()}. {item.ItemCode} - {item.TipoDocumento} - {item.NumeroDocumento} - {item.NumeroLote}");
                        res = false;
                        break;
                    }


                    //Descuento seguro: la resta SOLO se aplica si en ese instante
                    //todavía alcanza la cantidad en el lote de origen. Si dos
                    //usuarios chocan por el mismo lote, al segundo no le alcanza
                    //y se marca la fila como Faltante ("F") en vez de descontar
                    //de más y dejar el lote en negativo.
                    bool descontado = DescontarCantidadLoteSeguro(item.ItemCode, item.NumeroLote, item.Almacen, item.TipoDocumentoEntrada, item.NumeroDocumentoEntrada, item.Cantidad);
                    if (!descontado)
                    {
                        MarcarSalidaComoFaltante(item.TipoDocumento, item.NumeroDocumento, item.ItemCode, item.NumeroLote);
                        Console.WriteLine($"ATENCION: no alcanzó cantidad del lote {item.NumeroLote} para {item.ItemCode} (posible choque con otro usuario). Revisar CLC_LTSSA, quedó como Faltante (F).");
                        try { oApp.SetStatusBarMessage($"Atención: no alcanzó cantidad del lote {item.NumeroLote} para {item.ItemCode}. Revisar."); } catch { }
                    }

                }
                listaLotesSalida.Clear();
                return res;

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en crear lotes de articulos con lote entrada, {ex.Message}");
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(Lotes);
                Ini.ReleaseCOMObjetc(LotesEntrada);
                Ini.ReleaseCOMObjetc(oRs);
            }
        }

        /// <summary>
        /// Punto único de entrada para et_FORM_DATA_ADD de los documentos de salida.
        /// - Si es Borrador (pendiente de aprobación): guarda el registro como
        ///   Pendiente ("P"), SIN descontar cantidad del lote de origen todavía.
        /// - Si es el documento definitivo: si ya existe un pendiente para ese
        ///   Tipo+NumDoc (venía de un borrador aprobado, tal vez desde otra
        ///   sesión/usuario), descuenta la cantidad, recrea la entrada en el
        ///   almacen destino si es traslado (OWTR), y lo confirma. Si no existe
        ///   pendiente (flujo normal, sin aprobación), procesa directo como
        ///   Confirmado ("C") igual que siempre.
        /// </summary>
        internal bool ProcesarAdd(bool esBorrador, string tipo)
        {
            try
            {
                if (esBorrador)
                {
                    Console.WriteLine($"Documento {tipo} enviado a aprobación (Borrador). Salida de lote guardada como Pendiente.");
                    return SalidaLotes("P");
                }

                string numDoc = ObtenerNumDocFormularioActivo();

                if (!string.IsNullOrEmpty(numDoc) && ExistePendienteSalida(tipo, numDoc))
                {
                    Console.WriteLine($"Documento {tipo} {numDoc} venía de un borrador aprobado. Confirmando salida de lote pendiente.");
                    return ConfirmarPendientesSalida(tipo, numDoc);
                }

                return SalidaLotes("C");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en ProcesarAdd de articulos con lote salida, {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Lee el numero de documento (item "8") del formulario activo.
        /// </summary>
        private string ObtenerNumDocFormularioActivo()
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.EditText oDocNum = null;
            try
            {
                oForm = oApp.Forms.ActiveForm;
                oDocNum = (SAPbouiCOM.EditText)oForm.Items.Item("8").Specific;
                return oDocNum.String;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error obteniendo numero de documento activo, {ex.Message}");
                return string.Empty;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oDocNum);
                Ini.ReleaseCOMObjetc(oForm);
            }
        }

        /// <summary>
        /// Verifica si hay filas Pendientes en CLC_LTSSA para un Tipo+NumDoc de documento.
        /// </summary>
        private bool ExistePendienteSalida(string tipo, string numDoc)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = $"SELECT COUNT(*) AS \"Total\" FROM \"@CLC_LTSSA\" WHERE \"U_TipDoc\" = '{tipo}' AND \"U_NumDoc\" = {numDoc} AND \"U_Estado\" = 'P'";
                oRs.DoQuery(query);
                if (oRs.RecordCount > 0)
                {
                    oRs.MoveFirst();
                    int total = Convert.ToInt32(oRs.Fields.Item("Total").Value);
                    return total > 0;
                }
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error validando pendientes en CLC_LTSSA, {ex.Message}");
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oRs);
            }
        }

        /// <summary>
        /// Descuenta la cantidad del lote de origen para cada fila Pendiente de
        /// CLC_LTSSA que coincida con Tipo+NumDoc, y las pasa a Confirmado ("C").
        /// Si el tipo es OWTR (traslado) y el almacen destino maneja lotes,
        /// además recrea la entrada del lote en el almacen destino (igual que
        /// hace SalidaLotes() en el flujo normal sin aprobación).
        /// </summary>
        private bool ConfirmarPendientesSalida(string tipo, string numDoc)
        {
            SAPbobsCOM.Recordset oRs = null;
            SAPbobsCOM.Recordset oRsAux = null;
            SAPbobsCOM.UserTable LotesEntrada = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRsAux = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = $"SELECT \"Name\" AS \"ItemCode\", \"U_NumLot\", \"U_Cantidad\", \"U_almacen\", \"U_TDocEnt\", \"U_NDocEnt\", \"U_AlmDest\", \"U_FchProc\" " +
                               $"FROM \"@CLC_LTSSA\" WHERE \"U_TipDoc\" = '{tipo}' AND \"U_NumDoc\" = {numDoc} AND \"U_Estado\" = 'P'";
                oRs.DoQuery(query);

                if (oRs.RecordCount == 0)
                {
                    return true; // nada pendiente que confirmar
                }

                for (int i = 0; i < oRs.RecordCount; i++)
                {
                    string itemCode = oRs.Fields.Item("ItemCode").Value.ToString();
                    string numLot = oRs.Fields.Item("U_NumLot").Value.ToString();
                    double cantidad = Convert.ToDouble(oRs.Fields.Item("U_Cantidad").Value);
                    string almacen = oRs.Fields.Item("U_almacen").Value.ToString();
                    string tDocEnt = oRs.Fields.Item("U_TDocEnt").Value.ToString();
                    string nDocEnt = oRs.Fields.Item("U_NDocEnt").Value.ToString();
                    string almacenDest = oRs.Fields.Item("U_AlmDest").Value.ToString();
                    string fechaProceso = oRs.Fields.Item("U_FchProc").Value.ToString();

                    // 1) Descontar la cantidad del lote de origen.
                    // Descuento seguro: la resta SOLO se aplica si en ese instante
                    // todavía alcanza (protegido por la base de datos, evita choques
                    // entre usuarios concurrentes).
                    bool descontado = DescontarCantidadLoteSeguro(itemCode, numLot, almacen, tDocEnt, nDocEnt, cantidad);
                    if (!descontado)
                    {
                        MarcarSalidaComoFaltante(tipo, numDoc, itemCode, numLot);
                        Console.WriteLine($"ATENCION: al confirmar la aprobación, no alcanzó cantidad del lote {numLot} para {itemCode}. Quedó como Faltante (F).");
                    }

                    // 2) Si es traslado (OWTR) y el almacen destino maneja lotes,
                    //    recrear la entrada del lote en el almacen destino.
                    if (tipo.Equals("OWTR") && !string.IsNullOrEmpty(almacenDest))
                    {
                        string queryAlmacen = $"SELECT T1.\"WhsCode\" FROM OWHS T1 WHERE T1.\"WhsCode\" = '{almacenDest}' AND T1.\"U_UtilLot\" = 'Y'";
                        oRsAux.DoQuery(queryAlmacen);

                        if (oRsAux.RecordCount > 0)
                        {
                            oRsAux.MoveFirst();
                            string whsDestValido = oRsAux.Fields.Item(0).Value.ToString();

                            if (!string.IsNullOrEmpty(whsDestValido))
                            {
                                // CORRECCION: si el traslado tiene 2 líneas con el mismo
                                // artículo+lote hacia el mismo almacén destino, se SUMA la
                                // cantidad a la fila que ya existe en vez de duplicarla.
                                string queryBuscaDestino = $"SELECT \"Code\" FROM \"@CLC_LTS\" WHERE \"Name\" = '{itemCode}' " +
                                                            $"AND \"U_TipDoc\" = '{tipo}' AND \"U_NumDoc\" = {numDoc} AND \"U_NumLot\" = '{numLot}' " +
                                                            $"AND \"U_almacen\" = '{almacenDest}' AND \"U_Estado\" = 'C'";
                                oRsAux.DoQuery(queryBuscaDestino);

                                if (oRsAux.RecordCount > 0)
                                {
                                    oRsAux.MoveFirst();
                                    string codeExistenteDestino = oRsAux.Fields.Item("Code").Value.ToString();
                                    oRsAux.DoQuery($"UPDATE \"@CLC_LTS\" SET \"U_CantLote\" = (\"U_CantLote\" + {cantidad}) WHERE \"Code\" = '{codeExistenteDestino}'");
                                    Console.WriteLine($"Lote destino ya existía en el traslado aprobado, se sumó cantidad {itemCode} - {tipo} - {numDoc} - {numLot}");
                                }
                                else
                                {
                                    LotesEntrada = oCompany.UserTables.Item("CLC_LTS");
                                    LotesEntrada.Name = itemCode;
                                    LotesEntrada.UserFields.Fields.Item("U_TipDoc").Value = tipo;
                                    LotesEntrada.UserFields.Fields.Item("U_NumDoc").Value = numDoc;
                                    LotesEntrada.UserFields.Fields.Item("U_NumLot").Value = numLot;
                                    LotesEntrada.UserFields.Fields.Item("U_CantLote").Value = cantidad;
                                    LotesEntrada.UserFields.Fields.Item("U_FchVenLt").Value = fechaProceso;
                                    LotesEntrada.UserFields.Fields.Item("U_FchAdmLt").Value = DateTime.Now;
                                    LotesEntrada.UserFields.Fields.Item("U_almacen").Value = almacenDest;
                                    LotesEntrada.UserFields.Fields.Item("U_Estado").Value = "C";

                                    if (!LotesEntrada.Add().Equals(0))
                                    {
                                        Console.WriteLine($"Error al confirmar entrada de traslado con aprobación. {oCompany.GetLastErrorDescription()}. {itemCode} - {tipo} - {numDoc} - {numLot}");
                                        return false;
                                    }
                                    Ini.ReleaseCOMObjetc(LotesEntrada);
                                    LotesEntrada = null;
                                }
                            }
                        }
                    }

                    oRs.MoveNext();
                }

                string queryConfirma = $"UPDATE \"@CLC_LTSSA\" SET \"U_Estado\" = 'C' WHERE \"U_TipDoc\" = '{tipo}' AND \"U_NumDoc\" = {numDoc} AND \"U_Estado\" = 'P'";
                oRsAux.DoQuery(queryConfirma);

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error confirmando pendientes de salida en CLC_LTSSA, {ex.Message}");
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oRs);
                Ini.ReleaseCOMObjetc(oRsAux);
                Ini.ReleaseCOMObjetc(LotesEntrada);
            }
        }

        /// <summary>
        /// Descuenta la cantidad del lote de origen de forma SEGURA: la resta solo
        /// se aplica si en ese instante exacto todavía alcanza (la condición
        /// "U_CantLote >= cantidad" va dentro del mismo UPDATE, así que queda
        /// protegida por la base de datos). Si dos usuarios chocan por el mismo
        /// lote casi al mismo tiempo, al que llegue después esta función le
        /// devuelve false en vez de dejar el lote en negativo.
        /// </summary>
        private bool DescontarCantidadLoteSeguro(string itemCode, string numLot, string almacen, string tipDocEntrada, string numDocEntrada, double cantidad)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string filtro = $"\"Name\" = '{itemCode}' AND \"U_NumLot\" = '{numLot}' AND \"U_almacen\" = '{almacen}' " +
                                 $"AND \"U_TipDoc\" = '{tipDocEntrada}' AND \"U_NumDoc\" = {numDocEntrada}";

                // Cantidad ANTES de intentar el descuento (para poder comparar después).
                oRs.DoQuery($"SELECT \"U_CantLote\" FROM \"@CLC_LTS\" WHERE {filtro}");
                if (oRs.RecordCount == 0)
                {
                    Console.WriteLine($"No se encontró el lote de origen {numLot} de {itemCode} en {almacen} ({tipDocEntrada} {numDocEntrada}).");
                    return false;
                }
                oRs.MoveFirst();
                double cantidadAntes = Convert.ToDouble(oRs.Fields.Item("U_CantLote").Value);

                // Descuento protegido: SOLO se aplica si en este instante alcanza.
                oRs.DoQuery($"UPDATE \"@CLC_LTS\" SET \"U_CantLote\" = (\"U_CantLote\" - {cantidad}) WHERE {filtro} AND \"U_CantLote\" >= {cantidad}");

                // Cantidad DESPUES: si de verdad se aplicó, bajó exactamente "cantidad".
                oRs.DoQuery($"SELECT \"U_CantLote\" FROM \"@CLC_LTS\" WHERE {filtro}");
                oRs.MoveFirst();
                double cantidadDespues = Convert.ToDouble(oRs.Fields.Item("U_CantLote").Value);

                bool seAplico = Math.Abs((cantidadAntes - cantidadDespues) - cantidad) < 0.0001;
                return seAplico;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en descuento seguro de lote, {ex.Message}");
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oRs);
            }
        }

        /// <summary>
        /// Marca una fila de CLC_LTSSA como Faltante ("F") cuando, al momento de
        /// descontar, ya no alcanzaba la cantidad en el lote de origen (choque
        /// entre usuarios concurrentes). Queda ahí para revisión manual, en vez
        /// de aparecer como Confirmada ("C") sin haberse podido descontar de verdad.
        /// </summary>
        private void MarcarSalidaComoFaltante(string tipo, string numDoc, string itemCode, string numLot)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = $"UPDATE \"@CLC_LTSSA\" SET \"U_Estado\" = 'F' " +
                                $"WHERE \"U_TipDoc\" = '{tipo}' AND \"U_NumDoc\" = {numDoc} AND \"Name\" = '{itemCode}' AND \"U_NumLot\" = '{numLot}'";
                oRs.DoQuery(query);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error marcando salida como Faltante, {ex.Message}");
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oRs);
            }
        }

    }
}