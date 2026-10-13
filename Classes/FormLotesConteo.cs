using Lotes_Hornitos.Modelos;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Windows.Forms;

namespace Lotes_Hornitos.Classes
{
    internal class FormLotesConteo
    {
        private SAPbobsCOM.Company oCompany;
        private SAPbouiCOM.Application oApp;
        public static List<ListConteo> ListaConteo;
        public bool DesdeLista;

        public FormLotesConteo(SAPbobsCOM.Company oCompany, SAPbouiCOM.Application oApp)
        {
            this.oCompany = oCompany;
            this.oApp = oApp;
            ListaConteo = new List<ListConteo>();
        }

        /// <summary>
        /// Funcion para agregar al dt de crear la informacion de la linea
        /// </summary>
        /// <param name="pVal"></param>
        /// <param name="tipo"></param>
        internal void AddInfoLotes(ItemEvent pVal, string tipo)
        {
            SAPbouiCOM.Form oForm = null;
            Grid grid_Crear = null;
            Grid grid_LineasDoc = null;
            DataTable DT_Crear = null;
            DataTable DT_LineasDocumento = null;
            Grid grid_Sel = null;
            DataTable DT_Sel = null;

            SAPbouiCOM.CellPosition oCellPosition = null;
            EditTextColumn ColumnGrid = null;
            GridColumn oCol = null;
            ComboBoxColumn oComboCol = null;
            Columns columns = null;
            Column column = null;

            SAPbobsCOM.Recordset oRs = null;

            double cantidad = 0, cantidadDesviacion = 0;
            string articulo = string.Empty, query = string.Empty, queryLotesSalida = string.Empty;
            string path = string.Empty;


            var Sb = new System.Text.StringBuilder();
            try
            {
                oForm = oApp.Forms.Item(pVal.FormUID);
                oForm.Freeze(true);
                grid_Crear = (Grid)oForm.Items.Item("GrdCrear").Specific;
                grid_LineasDoc = (Grid)oForm.Items.Item("GrdLineas").Specific;
                DT_Crear = (DataTable)oForm.DataSources.DataTables.Item("DTCrea");
                DT_LineasDocumento = (DataTable)oForm.DataSources.DataTables.Item("DTLineas");
                oCellPosition = grid_LineasDoc.GetCellFocus();
                grid_Sel = (Grid)oForm.Items.Item("GrdSel").Specific;
                DT_Sel = (DataTable)oForm.DataSources.DataTables.Item("DTSel");

                cantidad = Convert.ToDouble(DT_LineasDocumento.GetValue("CantidadContada", pVal.Row));
                cantidadDesviacion = Convert.ToDouble(DT_LineasDocumento.GetValue("CantidadDesviacion", pVal.Row));

                // valida si hay informacion de lotes ya creados en la tabla de entradas, esto para mostrar la info de lotes en el gri de lotes creados

                //Cambio por Lista y datatable
                CargarLista(ref DT_Crear, ref DT_Sel, ListaConteo, DT_LineasDocumento.GetValue("ItemCode", pVal.Row).ToString(), DT_LineasDocumento.GetValue("Almacen", pVal.Row).ToString(), cantidad, DT_LineasDocumento.GetValue("Tipo Conteo", pVal.Row).ToString(), DT_LineasDocumento.GetValue("NumLinea", pVal.Row).ToString());


                //Si no hay lotes creados continua con la informacion de lotes a crear, desde la informacion del documento de entrada.
                if (DT_Crear.IsEmpty)
                {

                    Sb.Clear();
                    switch (DT_LineasDocumento.GetValue("Tipo Conteo", pVal.Row).ToString())
                    {
                        case "E":
                            query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.InfoArticulosconLote.sql")).ReadToEnd();
                            query = Sb.AppendFormat(query, DT_LineasDocumento.GetValue("ItemCode", pVal.Row), tipo, DT_LineasDocumento.GetValue("Numero Doc", pVal.Row), cantidad, DT_LineasDocumento.GetValue("Almacen", pVal.Row), DT_LineasDocumento.GetValue("NumLinea", pVal.Row)).ToString();
                            break;
                        case "S":
                            query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.InfoArticulosconLoteConteoSalida.sql")).ReadToEnd();
                            query = Sb.AppendFormat(query, DT_LineasDocumento.GetValue("ItemCode", pVal.Row), DT_LineasDocumento.GetValue("Almacen", pVal.Row), DT_LineasDocumento.GetValue("NumLinea", pVal.Row)).ToString();

                            break;
                    }



                    // carga la info en el grid de lotes a crear
                    DT_Crear.ExecuteQuery(query);
                }


                if (!DT_Crear.IsEmpty)
                {
                    grid_Crear.SelectionMode = BoMatrixSelect.ms_None;
                    grid_Crear.Item.Enabled = true;
                    grid_Crear.Item.AffectsFormMode = true;
                    grid_Crear.Columns.Item("ItemCode").Visible = false;
                    grid_Crear.Columns.Item("Tipo doc").Visible = false;
                    grid_Crear.Columns.Item("Num Doc").Visible = false;
                    grid_Crear.Columns.Item("Index").Visible = false;
                    grid_Crear.Columns.Item("NumLinea").Visible = false;

                    ColumnGrid = (EditTextColumn)grid_Crear.Columns.Item("Sel");
                    ColumnGrid.Type = BoGridColumnType.gct_CheckBox;

                    oForm.Items.Item("GrdSel").Enabled = false;
                    oForm.Items.Item("BtnAgre").Enabled = false;
                    oForm.Items.Item("BtnElim").Enabled = false;
                    oForm.Items.Item("BtAddln").Enabled = false;
                    oForm.Items.Item("BtDelLn").Enabled = false;

                    for (int i = 0; i < grid_Crear.Columns.Count; i++)
                    {

                        grid_Crear.Columns.Item(i).Editable = DesdeLista ? false : true;
                        grid_Crear.Columns.Item(i).AffectsFormMode = DesdeLista ? false : true;
                    }

                    grid_Crear.Columns.Item("Sel").Editable = true;

                    grid_Crear.AutoResizeColumns();

                    grid_Crear.Columns.Item("Cantidad").Editable = true;


                    //if (Convert.ToDouble( DT_Crear.GetValue("Cantidad", grid_Crear.Rows.Count - 1)) < 0)
                    //{
                    //    grid_Crear.Columns.Item("Numero lote").Type =SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                    //    oComboCol = (SAPbouiCOM.ComboBoxColumn)oCol;
                    //    oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //    Sb.Clear();
                    //    queryLotesSalida = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.LotesSalidaConteo.sql")).ReadToEnd();
                    //    queryLotesSalida = Sb.AppendFormat(queryLotesSalida, DT_LineasDocumento.GetValue("ItemCode", pVal.Row), DT_LineasDocumento.GetValue("Almacen", pVal.Row),Math.Abs(cantidad)).ToString();
                    //    oRs.DoQuery(queryLotesSalida);

                    //}

                    var s = DT_Crear.GetValue("Fecha vence", grid_Crear.Rows.Count - 1);
                    if (DT_Crear.GetValue("Fecha vence", grid_Crear.Rows.Count - 1).ToString().Contains("1900"))
                    {
                        DT_Crear.SetValue("Fecha vence", grid_Crear.Rows.Count - 1, "");
                    }
                    if (cantidad < 0)
                    {
                        grid_Crear.Columns.Item("Fecha Admin").TitleObject.Caption = "Fecha salida";
                    }

                    //oForm.Items.Item("BtnAddLn").Enabled = true;
                    if (DT_LineasDocumento.GetValue("Tipo Conteo", pVal.Row).ToString().Equals("S"))
                    {
                        grid_Crear.Columns.Item("Numero lote").Editable = false;
                        grid_Crear.Columns.Item("Cantidad disponible").Editable = false;
                        grid_Crear.Columns.Item("Cantidad").Editable = true;
                        grid_Crear.Columns.Item("Fecha vence").Editable = false;
                        grid_Crear.Columns.Item("Fecha Admin").Editable = false;
                        grid_Crear.Columns.Item("Almacen").Editable = false;
                        oForm.Items.Item("GrdSel").Enabled = true;
                        oForm.Items.Item("BtnAgre").Enabled = true;
                        oForm.Items.Item("BtnElim").Enabled = true;

                    }
                    else
                    {
                        oForm.Items.Item("BtAddln").Enabled = true;
                        oForm.Items.Item("BtDelLn").Enabled = true;
                    }

                }
                else
                {
                    grid_Crear.SelectionMode = BoMatrixSelect.ms_None;
                    grid_Crear.Item.Enabled = false;
                }

                if (DT_LineasDocumento.GetValue("Tipo Conteo", pVal.Row).ToString().Equals("S"))
                {
                    if (!DT_Sel.IsEmpty)
                    {
                        grid_Sel.Columns.Item("Sel").Visible = false;
                        grid_Sel.Columns.Item("Tipo doc").Visible = false;
                        grid_Sel.Columns.Item("Num Doc").Visible = false;
                        grid_Sel.Columns.Item("Index").Visible = false;
                        grid_Sel.Columns.Item("NumLinea").Visible = false;
                    }

                }

                //columns = (SAPbouiCOM.Columns)grid_Crear.Columns;
                //column = columns.Item("Fecha vence");
                //column.DataBind.SetBound(true,"@CLC_LTS","U_FchVenLt");





            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
            }
            finally
            {
                oForm.Freeze(false);
                Ini.ReleaseCOMObjetc(oForm);
                Ini.ReleaseCOMObjetc(grid_Crear);
                Ini.ReleaseCOMObjetc(grid_LineasDoc);
                Ini.ReleaseCOMObjetc(DT_Crear);
                Ini.ReleaseCOMObjetc(DT_LineasDocumento);
                Ini.ReleaseCOMObjetc(oCellPosition);
                Ini.ReleaseCOMObjetc(ColumnGrid);
                Ini.ReleaseCOMObjetc(columns);
                Ini.ReleaseCOMObjetc(column);
                Ini.ReleaseCOMObjetc(oCol);
                Ini.ReleaseCOMObjetc(oComboCol);
                Ini.ReleaseCOMObjetc(oRs);

            }
        }

        private void CargarLista(ref DataTable dT_Crear, ref DataTable dT_Sel, List<ListConteo> listaConteo, string _ItemCode, string almacen, double cantidadlinea, string tipoConteo, string RowNum)
        {
            string query = string.Empty, queryConteo = string.Empty;
            int cont = 0;
            double cantidadEntrada = 0;
            List<string> ListLotes = new List<string>();
            var Sb = new System.Text.StringBuilder();

            dT_Crear.Rows.Clear();
            dT_Sel.Rows.Clear();

            if (listaConteo == null || !listaConteo.Any())
            {
                DesdeLista = false;
                return;
            }


            query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.LotesDatatable.sql")).ReadToEnd();


            for (int i = 0; i < listaConteo.Count; i++)
            {
                var item = listaConteo[i];

                if (item.ItemCode.Equals(_ItemCode))
                {
                    if (item.almacen.Equals(almacen))
                    {
                        ListLotes.Add(Sb.AppendFormat(query, item.ItemCode, item.NumDoc, item.TipDoc, item.CantLote, item.NumLot, item.FchVenLt, item.FchAdmLt, item.almacen, i, item.NumLinea).ToString());
                        cantidadEntrada += item.CantLote;
                        Sb.Clear();
                    }

                }



            }

            if (ListLotes.Any())
            {
                query = string.Join(" UNION ALL ", ListLotes);
                if (tipoConteo.Equals("E"))
                {
                    dT_Crear.ExecuteQuery(query);
                    if (cantidadlinea > cantidadEntrada)
                    {
                        DesdeLista = false;

                    }
                    else
                    {
                        DesdeLista = true;

                    }
                }
                else
                {
                    queryConteo = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.InfoArticulosconLoteConteoSalida.sql")).ReadToEnd();
                    queryConteo = Sb.AppendFormat(queryConteo, _ItemCode, almacen, RowNum).ToString();
                    Sb.Clear();
                    dT_Sel.ExecuteQuery(query);
                    dT_Crear.ExecuteQuery(queryConteo);

                    if (cantidadlinea > cantidadEntrada)
                    {
                        DesdeLista = false;

                    }
                    else
                    {
                        DesdeLista = true;

                    }
                }


            }
            else
            {
                DesdeLista = false;
            }

        }

        internal void CleanData(ItemEvent pVal)
        {
            SAPbouiCOM.Form oForm = null;
            Grid grid_Crear = null;
            Grid grid_LineasDoc = null;
            DataTable DT_Crear = null;
            DataTable DT_LineasDocumento = null;
            Grid grid_Sel = null;
            DataTable DT_Sel = null;

            try
            {
                oForm = oApp.Forms.Item(pVal.FormUID);
                oForm.Freeze(true);
                grid_Crear = (Grid)oForm.Items.Item("GrdCrear").Specific;
                grid_LineasDoc = (Grid)oForm.Items.Item("GrdLineas").Specific;
                DT_Crear = (DataTable)oForm.DataSources.DataTables.Item("DTCrea");
                DT_LineasDocumento = (DataTable)oForm.DataSources.DataTables.Item("DTLineas");
                DT_Sel = (DataTable)oForm.DataSources.DataTables.Item("DTSel");
                grid_Sel = (Grid)oForm.Items.Item("GrdSel").Specific;


                for (int i = 0; i < DT_LineasDocumento.Rows.Count; i++)
                {
                    if (DT_LineasDocumento.GetValue("Sel", i).Equals("Y"))
                    {
                        DT_LineasDocumento.SetValue("Sel", i, "N");
                    }
                }

                DT_Crear.Clear();
                DT_Sel.Clear();
                oForm.Freeze(false);


            }
            catch (Exception ex)
            {
                oForm.Freeze(false);

            }
            finally
            {
                oForm.Freeze(false);

                Ini.ReleaseCOMObjetc(oForm);
                Ini.ReleaseCOMObjetc(grid_Crear);
                Ini.ReleaseCOMObjetc(grid_LineasDoc);
                Ini.ReleaseCOMObjetc(DT_Crear);
                Ini.ReleaseCOMObjetc(DT_LineasDocumento);
                Ini.ReleaseCOMObjetc(DT_Sel);
                Ini.ReleaseCOMObjetc(grid_Sel);
            }
        }

        internal bool AddLotesList(ItemEvent pVal)
        {
            SAPbouiCOM.Form oForm = null;
            Grid grid_Crear = null;
            Grid grid_Lineas = null;
            DataTable DT_Crear = null;
            DataTable DT_Lineas = null;
            Grid grid_Seleccion = null;
            DataTable DT_Seleccion = null;

            object fechaVence = null;
            Console.WriteLine("Inicio de proceso de creación lote");
            bool result = true;

            string loteAIngresar = string.Empty, LoteDisponibleLinea = string.Empty, LoteSeleccionadoLinea = string.Empty; ;
            int rowDtLines = 0;
            double cantidadAingresar = 0, cantidadLineas = 0, cantidadDisponible = 0, cantidadDisponibleSelecionada = 0, cantidadDisponibleLinea = 0, cantidadDisponibleSelecionadaLinea = 0;
            try
            {
                oForm = oApp.Forms.Item(pVal.FormUID);
                grid_Crear = (Grid)oForm.Items.Item("GrdCrear").Specific;
                grid_Lineas = (Grid)oForm.Items.Item("GrdLineas").Specific;
                DT_Crear = (DataTable)oForm.DataSources.DataTables.Item("DTCrea");
                DT_Lineas = (DataTable)oForm.DataSources.DataTables.Item("DTLineas");
                grid_Seleccion = (Grid)oForm.Items.Item("GrdSel").Specific;
                DT_Seleccion = (DataTable)oForm.DataSources.DataTables.Item("DTSel");



                if (oForm.Mode.Equals(BoFormMode.fm_UPDATE_MODE))
                {
                    for (int i = 0; i < DT_Lineas.Rows.Count; i++)
                    {
                        if (DT_Lineas.GetValue("Sel", i).Equals("Y"))
                        {
                            if (DT_Lineas.GetValue("Tipo Conteo", i).ToString().Equals("E"))
                            {
                                for (int j = 0; j <= DT_Crear.Rows.Count - 1; j++)
                                {
                                    loteAIngresar = DT_Crear.GetValue("Numero lote", j).ToString();
                                    var val = Convert.ToDouble(DT_Crear.GetValue("Cantidad", j));


                                    if (string.IsNullOrEmpty(loteAIngresar))
                                    {
                                        oApp.MessageBox("Debe ingresar un número lote valido", 1, "Ok");
                                        result = false;
                                        break;
                                    }
                                    else if (Convert.ToDouble(DT_Crear.GetValue("Cantidad", j)) > 0)
                                    {
                                        if (!ValidaLoteCrear(loteAIngresar, DT_Crear.GetValue("Almacen", j).ToString(), DT_Crear.GetValue("ItemCode", j).ToString()))
                                        {
                                            oApp.MessageBox("El lote ingresado ya existe", 1, "Ok");
                                            result = false;
                                            break;
                                        }
                                    }


                                    //valida que no haya lotes iguales dentro del grid
                                    for (int k = j + 1; k <= DT_Crear.Rows.Count - 1; k++)
                                    {
                                        if (DT_Crear.GetValue("Numero lote", k).ToString().Equals(loteAIngresar))
                                        {
                                            oApp.MessageBox("Hay lotes duplicados", 1, "Ok");
                                            result = false;
                                            break;
                                        }
                                    }


                                    fechaVence = DT_Crear.GetValue("Fecha vence", j);

                                    if (fechaVence is null || string.IsNullOrEmpty(fechaVence.ToString()))
                                    {
                                        oApp.MessageBox("Debe ingresar una fecha vence valida", 1, "Ok");
                                        result = false;
                                        break;
                                    }

                                    if (!String.IsNullOrEmpty(DT_Crear.GetValue("Cantidad", j).ToString()))
                                    {
                                        cantidadAingresar += Convert.ToDouble(DT_Crear.GetValue("Cantidad", j));
                                    }


                                }

                                if (!result) { return result; }




                                for (int j = 0; j <= grid_Lineas.Rows.Count - 1; j++)
                                {
                                    if (DT_Lineas.GetValue("Sel", j).ToString().Equals("Y"))
                                    {
                                        cantidadLineas = Convert.ToDouble(DT_Lineas.GetValue("CantidadContada", j));
                                        rowDtLines = j;
                                        break;
                                    }

                                }

                                if (cantidadAingresar > cantidadLineas)
                                {
                                    oApp.MessageBox("La cantidad del lote no debe ser mayor a la cantidad del documento", 1, "Ok");
                                    result = false;
                                    return result;
                                }

                                // aca debe ir la asignacion de los lotes a el datatable global.

                                ConvertirALista(DT_Crear, DT_Lineas.GetValue("Numero Doc", rowDtLines).ToString(), DT_Lineas.GetValue("Tipo Doc", rowDtLines).ToString(), DT_Lineas.GetValue("Tipo Conteo", i).ToString());


                            }
                            else
                            {
                                //Capturar el valor principal de cantidad de la salida
                                cantidadLineas = Math.Abs(Convert.ToDouble(DT_Lineas.GetValue("CantidadContada", i)));

                                //Capturar el valor de cantidad de lote ingresado como disponible en seleccionados
                                for (int j = 0; j <= grid_Seleccion.Rows.Count - 1; j++)
                                {

                                    cantidadDisponibleSelecionada += Convert.ToDouble(DT_Seleccion.GetValue("Cantidad", j));

                                }

                                //Capturar el valor de cantidad de lote ingresado como disponible en disponibles
                                for (int j = 0; j <= grid_Crear.Rows.Count - 1; j++)
                                {
                                    if (DT_Crear.GetValue("Sel", j).ToString().Equals("Y"))
                                    {
                                        cantidadDisponible += Convert.ToDouble(DT_Crear.GetValue("Cantidad", j));
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
                                for (int j = 0; j <= grid_Crear.Rows.Count - 1; j++)
                                {
                                    if (DT_Crear.GetValue("Sel", j).ToString().Equals("Y"))
                                    {
                                        cantidadDisponibleLinea = Convert.ToDouble(DT_Crear.GetValue("Cantidad", j));
                                        LoteDisponibleLinea = DT_Crear.GetValue("Numero lote", j).ToString();

                                        for (int k = 0; k <= grid_Seleccion.Rows.Count - 1; k++)
                                        {
                                            var numlinea = (int)DT_Seleccion.GetValue("NumLinea", k);
                                            if (numlinea.Equals(j))
                                            {
                                                cantidadDisponibleSelecionadaLinea = Convert.ToDouble(DT_Seleccion.GetValue("Cantidad", k));
                                                LoteSeleccionadoLinea = DT_Seleccion.GetValue("Numero lote", k).ToString();
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


                                AddLotesSalidaConteoToLits(DT_Seleccion, grid_Lineas, DT_Lineas, DT_Lineas.GetValue("Tipo Conteo", i).ToString());
                            }
                        }

                    }
                    //if (DT_Lineas.GetValue("Tipo Conteo", pVal.Row).ToString().Equals("S"))
                    //{

                    //}


                }




                return result;
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oForm);
                Ini.ReleaseCOMObjetc(grid_Crear);
                Ini.ReleaseCOMObjetc(DT_Lineas);
                Ini.ReleaseCOMObjetc(DT_Crear);
                Ini.ReleaseCOMObjetc(grid_Lineas);
                Ini.ReleaseCOMObjetc(DT_Seleccion);
                Ini.ReleaseCOMObjetc(grid_Seleccion);
            }
        }

        /// <summary>
        /// Agregar a la lista los lotes de salida en el conteo
        /// </summary>
        /// <param name="dT_Seleccion">Datatable SAP de lotes selwccionados</param>
        /// <param name="grid_Lineas">Grid SAP de las lineas del docuemento</param>
        /// <param name="dT_Lineas">Datatable de las lineas del documento</param>
        private void AddLotesSalidaConteoToLits(DataTable dT_Seleccion, Grid grid_Lineas, DataTable dT_Lineas, string tipoconteo)
        {
            var tipo = Ini.Tipo;


            for (int i = 0; i <= grid_Lineas.Rows.Count - 1; i++)
            {
                if (dT_Lineas.GetValue("Sel", i).ToString().Equals("Y"))
                {
                    for (int j = 0; j < dT_Seleccion.Rows.Count; j++)
                    {


                        var fechavence = dT_Seleccion.GetValue("Fecha vence", j);
                        DateTime fecha = Convert.ToDateTime(fechavence);

                        var fechaAdmin = dT_Seleccion.GetValue("Fecha Admin", j);
                        DateTime fechaA = Convert.ToDateTime(fechaAdmin);

                        var lote = new ListConteo
                        {
                            ItemCode = dT_Seleccion.GetValue("ItemCode", j).ToString(),
                            NumLot = dT_Seleccion.GetValue("Numero lote", j).ToString(),
                            almacen = dT_Lineas.GetValue("Almacen", i).ToString(),
                            NumDoc = (int)dT_Lineas.GetValue("Numero Doc", i),
                            TipDoc = dT_Lineas.GetValue("Tipo Doc", i).ToString(),
                            CantLote = Convert.ToDouble(dT_Seleccion.GetValue("Cantidad", j).ToString()),
                            FchVenLt = fecha.ToString("dd/MM/yyyy"),
                            FchAdmLt = fechaA.ToString("dd/MM/yyyy"),
                            NumLinea = (int)dT_Lineas.GetValue("NumLinea", i),
                            NumeroDocumentoEntrada = dT_Seleccion.GetValue("Num Doc", j).ToString(),
                            TipoDocumentoEntrada = dT_Seleccion.GetValue("Tipo doc", j).ToString(),
                            TipoConteo = tipoconteo
                        };



                        ListaConteo.Add(lote);

                    }
                }
            }
        }

        public void ConvertirALista(SAPbouiCOM.DataTable dtSAP, string NumDocConteo, string TipoDocConteo, string TipoConteo)
        {
            //var lista = new List<EntradaLotes>();

            for (int i = 0; i < dtSAP.Rows.Count; i++)
            {

                //add
                var item = new ListConteo();

                item.ItemCode = dtSAP.GetValue("ItemCode", i)?.ToString();
                item.TipDoc = dtSAP.GetValue("Tipo doc", i)?.ToString();
                item.NumDoc = (int)dtSAP.GetValue("Num Doc", i);
                item.NumLot = dtSAP.GetValue("Numero lote", i)?.ToString();
                item.CantLote = Convert.ToDouble(dtSAP.GetValue("Cantidad", i));

                var fechaObj = dtSAP.GetValue("Fecha vence", i);
                DateTime fecha = Convert.ToDateTime(fechaObj);
                item.FchVenLt = fecha.ToString("dd/MM/yyyy");

                //if (
                //    DateTime.TryParse(fechaObj.ToString(), out DateTime fecha))
                //{
                //    item.FchVenLt = fecha;
                //}

                fechaObj = dtSAP.GetValue("Fecha Admin", i);
                fecha = Convert.ToDateTime(fechaObj);
                item.FchAdmLt = fecha.ToString("dd/MM/yyyy");

                item.almacen = dtSAP.GetValue("Almacen", i)?.ToString();
                item.NumLinea = (int)dtSAP.GetValue("NumLinea", i);

                item.NumeroDocumentoEntrada = NumDocConteo;
                item.TipoDocumentoEntrada = TipoDocConteo;
                item.TipoConteo = TipoConteo;


                bool existe = ListaConteo.Any(x =>
                x.ItemCode == item.ItemCode &&
                x.NumDoc == item.NumDoc &&
                x.NumLot == item.NumLot &&
                x.NumLinea == item.NumLinea
                );

                if (!existe)
                {
                    ListaConteo.Add(item);
                }
            }

            //return lista;
        }

        private bool ValidaLoteCrear(string loteAIngresar, string almacen, string ItemCode)
        {
            SAPbobsCOM.Recordset oRs = null;
            string query = string.Empty;

            try
            {
                query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ValidaLoteCrear.sql")).ReadToEnd();
                query = query.Replace("lote", loteAIngresar).Replace("bodega", almacen).Replace("item", ItemCode);

                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRs.DoQuery(query);

                if (oRs == null)
                {
                    return false;
                }
                else
                {
                    oRs.MoveFirst();
                    if ((int)oRs.Fields.Item(0).Value > 0)
                    {
                        if (oApp.MessageBox($"Para el articulo {ItemCode} ya existe el lote {loteAIngresar} en el almacen {almacen}.\n ¿Desea continuar?", 1, "Si", "No").Equals(1))
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }


                    }
                    else
                    {
                        return true;
                    }
                }
            }
            catch
            {
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oRs);
            }

        }

        internal List<bool> ValidasiexisteloteentablaConteo(string query)
        {
            double cantidadLote = 0;
            bool existelote = false, existeenconsulta = false;
            SAPbobsCOM.Recordset oRs = null;

            List<bool> result = new List<bool>();

            try
            {

                if (ListaConteo == null || !ListaConteo.Any())
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

                    for (int j = 0; j < ListaConteo.Count; j++)
                    {
                        var item = ListaConteo[j];

                        System.Console.WriteLine($"desde lista {item.ItemCode}");
                        System.Console.WriteLine($"desde recordset {oRs.Fields.Item("ItemCode").Value.ToString()}");

                        if (item.NumLinea.Equals((int)oRs.Fields.Item("NumLinea").Value))
                        {
                            if (item.ItemCode.Equals(oRs.Fields.Item("ItemCode").Value.ToString()))
                            {

                                if (!item.almacen.Equals(oRs.Fields.Item("Almacen").Value.ToString()))
                                {

                                    ListaConteo.RemoveAt(j);
                                    result.Add(false);
                                    break;
                                }

                                cantidadLote += item.CantLote;
                                existelote = true;
                            }
                        }




                    }
                    if (existelote)
                    {
                        if (!cantidadLote.Equals(Math.Abs(Convert.ToDouble(oRs.Fields.Item("Cantidad").Value))))
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
                for (o = 0; o < ListaConteo.Count; o++)
                {
                    existeenconsulta = false;
                    var item = ListaConteo[o];
                    oRs.MoveFirst();
                    for (int i = 1; i <= oRs.RecordCount; i++)
                    {
                        if (item.NumLinea.Equals((int)oRs.Fields.Item("NumLinea").Value))
                        {
                            if (item.ItemCode.Equals(oRs.Fields.Item("ItemCode").Value.ToString()))
                            {
                                if (item.almacen.Equals(oRs.Fields.Item("Almacen").Value.ToString()))
                                {
                                    //if (item.CantLote.Equals(Math.Abs(Convert.ToDouble(oRs.Fields.Item("Cantidad").Value))))
                                    //{
                                    existeenconsulta = true;
                                    //}
                                }
                            }
                        }
                        oRs.MoveNext();
                    }
                    if (!existeenconsulta)
                    {
                        ListaConteo.Remove(item);
                        o--;
                    }

                }





                //oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                //query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ArticulosenTablaEntrada.sql")).ReadToEnd();
                //query = Sb.AppendFormat(query, ItemCode, tipo, oDocNum).ToString();

                //oRs.DoQuery(query);
                //if (oRs == null)
                //{
                //    return false;
                //}
                //else
                //{
                //    oRs.MoveFirst();

                //    int valor = (int)oRs.Fields.Item(0).Value;
                //    if (valor.Equals(0)) return false;
                //}


                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en ValidasiexisteloteentablaEntrada {ex.Message}");
                return result;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oRs);
            }
        }

        public bool MostarItems(ItemEvent pVal, string Tipo)
        {
            string query = string.Empty;
            string path = string.Empty;

            SAPbouiCOM.Form Entrada = null;
            SAPbouiCOM.Matrix LineasEntrada = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.EditText oCantidadContada = null;
            SAPbouiCOM.EditText oCantidadDesviacion = null;
            SAPbouiCOM.EditText oAlmacen = null;
            SAPbouiCOM.EditText oDocNum = null;
            SAPbouiCOM.EditText RowNumber = null;


            SAPbouiCOM.EditText BaseType = null;


            double cantidadContada = 0, cantidadDesviacion = 0;
            //string cantidad = string.Empty;
            bool result = false;
            List<string> Items = new List<string>();
            var Sb = new System.Text.StringBuilder();

            try
            {
                Console.WriteLine("Inicio de consulta de articulos con lotes");

                Entrada = oApp.Forms.Item(pVal.FormUID);

                if (!Entrada.Mode.Equals(BoFormMode.fm_ADD_MODE)) return true;

                LineasEntrada = (SAPbouiCOM.Matrix)Entrada.Items.Item("1470000021").Specific;


                query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ArticulosconLoteConteo.sql")).ReadToEnd();

                for (int i = 1; i <= LineasEntrada.RowCount; i++)
                {
                    oEditText = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("1470000001").Cells.Item(i).Specific;



                    if (!string.IsNullOrEmpty(oEditText.String))
                    {

                        oCantidadDesviacion = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("1470000031").Cells.Item(i).Specific;
                        oCantidadContada = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("1470000013").Cells.Item(i).Specific;

                        oAlmacen = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("1470000007").Cells.Item(i).Specific;

                        oDocNum = (SAPbouiCOM.EditText)Entrada.Items.Item("1470000016").Specific;

                        RowNumber = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("1470000000").Cells.Item(i).Specific;

                        cantidadContada = Ini.ConvertToSystemFormat(oCantidadContada.String);
                        cantidadDesviacion = Ini.ConvertToSystemFormat(oCantidadDesviacion.String);
                        Items.Add(Sb.AppendFormat(query, oEditText.String, oDocNum.String, oAlmacen.String, cantidadDesviacion, Tipo, RowNumber.Value, cantidadContada).ToString());
                        Sb.Clear();
                        System.Console.WriteLine(oEditText.String);

                    }
                }



                //query = string.Join(",", Items.Select(x => $"'{x}'"));
                query = string.Join(" UNION ALL ", Items);

                Console.WriteLine("fin de consulta de articulos con lotes");

                // Mostar el formulario de lotes.
                System.Console.WriteLine("Mostrar el formulario de entrada de lotes");
                result = showForm(query, Tipo);
                System.Console.WriteLine("Fin carga de formulario de entrada de lotes");

                return result;

            }
            catch (Exception ex)
            {
                Entrada.Freeze(false);
                oApp.SetStatusBarMessage("Error Add-On " + Ini.asm.GetName().Name + " : " + ex.Message.ToString() + " " + ex.StackTrace.ToString());
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(Entrada);
                Ini.ReleaseCOMObjetc(LineasEntrada);
                Ini.ReleaseCOMObjetc(oEditText);
                Ini.ReleaseCOMObjetc(oCantidadContada);
                Ini.ReleaseCOMObjetc(oCantidadDesviacion);
                Ini.ReleaseCOMObjetc(oAlmacen);
                Ini.ReleaseCOMObjetc(oDocNum);
                Ini.ReleaseCOMObjetc(RowNumber);
            }

        }

        internal bool showForm(string query, string tipo)
        {
            System.Xml.XmlDocument oXmlDocument = null;

            string stringForm = null, oPath = null, FormUID = null;

            SAPbouiCOM.Form oForm = null;
            Grid grid_Crear = null;
            Grid grid_LineasDoc = null;
            DataTable DT_Crear = null;
            DataTable DT_LineasDocumento = null;
            EditTextColumn ColumnGrid = null;
            List<bool> result = new List<bool>();
            bool res = true;
            try
            {

                //validar si ya se agregaron los lotes

                result = ValidasiexisteloteentablaConteo(query);

                foreach (bool l in result)
                {
                    if (!l)
                    {
                        res = l;
                        break;
                    }
                }

                if (res)
                {
                    return true;
                }

                oXmlDocument = new System.Xml.XmlDocument();


                oPath = $"{Assembly.GetExecutingAssembly().GetName().Name}.Resources.XML.FRMLCON.srf";

                Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(oPath);
                StreamReader sr = new StreamReader(stream);
                stringForm = sr.ReadToEnd();

                oXmlDocument.LoadXml(stringForm);
                FormUID = oXmlDocument.SelectSingleNode("/application/forms/action/form | /Application/forms/action/form").Attributes["uid"].Value;
                oApp.LoadBatchActions(ref stringForm);

                oForm = oApp.Forms.Item(FormUID);
                grid_Crear = (Grid)oForm.Items.Item("GrdCrear").Specific;
                grid_LineasDoc = (Grid)oForm.Items.Item("GrdLineas").Specific;
                DT_Crear = (DataTable)oForm.DataSources.DataTables.Item("DTCrea");
                DT_LineasDocumento = (DataTable)oForm.DataSources.DataTables.Item("DTLineas");

                //DT_Crear.Rows.Add();




                //query = "SELECT 'N' AS 'Sel',15 AS 'Numero Doc',T0.ItemCode,CAST((1.000000) AS INT) AS 'Cantidad','01' AS 'Almacen'FROM OITM T0 WHERE T0.ItemCode = 'TESTLOTES-H2' AND T0.QryGroup28 = 'Y'";

                DT_LineasDocumento.ExecuteQuery(query);



                if (!DT_LineasDocumento.IsEmpty)
                {

                    grid_LineasDoc.SelectionMode = BoMatrixSelect.ms_None;
                    grid_LineasDoc.Item.Enabled = true;
                    grid_LineasDoc.Item.AffectsFormMode = false;


                    ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("ItemCode");
                    ColumnGrid.LinkedObjectType = "4";
                    ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("Almacen");
                    ColumnGrid.LinkedObjectType = "64";
                    ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("Sel");
                    ColumnGrid.Type = BoGridColumnType.gct_CheckBox;


                    for (int i = 0; i < grid_LineasDoc.Columns.Count; i++)
                    {

                        grid_LineasDoc.Columns.Item(i).Editable = false;
                        grid_LineasDoc.Columns.Item(i).TitleObject.Sortable = false;
                        grid_LineasDoc.Columns.Item(i).AffectsFormMode = false;
                    }

                    grid_LineasDoc.Columns.Item("Sel").Editable = true;
                    grid_LineasDoc.Columns.Item("Tipo Doc").Visible = false;
                    grid_LineasDoc.Columns.Item("Tipo Conteo").Visible = false;
                    grid_LineasDoc.Columns.Item("CantidadDesviacion").Visible = false;
                    grid_LineasDoc.Columns.Item("ItemCode").TitleObject.Caption = "Número de articulo";
                    grid_LineasDoc.Columns.Item("CantidadContada").TitleObject.Caption = "Cantidad Contada";
                    //grid_LineasDoc.Columns.Item("Cantidad solicitada").TitleObject.Caption = "Cantidad";
                    grid_LineasDoc.Columns.Item("Numero Doc").TitleObject.Caption = "Número documento";
                    grid_LineasDoc.AutoResizeColumns();
                    oForm.Visible = true;

                }
                else
                {
                    grid_LineasDoc.SelectionMode = BoMatrixSelect.ms_None;
                    grid_LineasDoc.Item.Enabled = false;
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
                Ini.ReleaseCOMObjetc(oForm);
                Ini.ReleaseCOMObjetc(grid_Crear);
                Ini.ReleaseCOMObjetc(grid_LineasDoc);
                Ini.ReleaseCOMObjetc(DT_Crear);
                Ini.ReleaseCOMObjetc(DT_LineasDocumento);
                Ini.ReleaseCOMObjetc(ColumnGrid);

            }
        }

        public bool CrearLotes()
        {
            bool res = true;
            SAPbobsCOM.UserTable Lotes = null;
            List<string> ListaUpdateEntrada = new List<string>();
            var Sb = new System.Text.StringBuilder();
            var query = string.Empty;

            try
            {
                for (int i = 0; i < ListaConteo.Count; i++)
                {
                    var item = ListaConteo[i];


                    Lotes = oCompany.UserTables.Item("CLC_LTSSATEMP");

                    Lotes.Name = item.ItemCode;
                    Lotes.UserFields.Fields.Item("U_TipDoc").Value = item.TipDoc;
                    Lotes.UserFields.Fields.Item("U_NumDoc").Value = item.NumDoc;
                    Lotes.UserFields.Fields.Item("U_NumLot").Value = item.NumLot;
                    Lotes.UserFields.Fields.Item("U_CantLote").Value = item.CantLote;
                    Lotes.UserFields.Fields.Item("U_FchVenLt").Value = item.FchVenLt;
                    Lotes.UserFields.Fields.Item("U_FchAdmLt").Value = item.FchAdmLt;
                    Lotes.UserFields.Fields.Item("U_almacen").Value = item.almacen;
                    Lotes.UserFields.Fields.Item("U_TipoCont").Value = item.TipoConteo;
                    Lotes.UserFields.Fields.Item("U_TpDocEnt").Value = item.TipoDocumentoEntrada;
                    Lotes.UserFields.Fields.Item("U_NmDocEnt").Value = item.NumeroDocumentoEntrada;
                    Lotes.UserFields.Fields.Item("U_RowNum").Value = item.NumLinea;

                    if (Lotes.Add().Equals(0))
                    {
                        Console.WriteLine($"Datos agregados a tabla entradas lotes exitosamente {item.ItemCode} - {item.TipDoc} - {item.NumDoc} - {item.NumLot}");

                    }
                    else
                    {
                        Console.WriteLine($"Error al agregar datos a tabla entrada de lotes. {oCompany.GetLastErrorDescription()}. {item.ItemCode} - {item.TipDoc} - {item.NumDoc} - {item.NumLot}");
                        res = false;
                        break;
                    }

                    //Lotes = oCompany.UserTables.Item("CLC_LTSSA");

                    //Lotes.Name = item.ItemCode;
                    //Lotes.UserFields.Fields.Item("U_TipDoc").Value = item.TipDoc;
                    //Lotes.UserFields.Fields.Item("U_NumDoc").Value = item.NumDoc;
                    //Lotes.UserFields.Fields.Item("U_NumLot").Value = item.NumLot;
                    //Lotes.UserFields.Fields.Item("U_Cantidad").Value = Math.Abs(item.CantLote);
                    //Lotes.UserFields.Fields.Item("U_FchSal").Value = DateTime.Now;
                    //Lotes.UserFields.Fields.Item("U_almacen").Value = item.almacen;

                    //if (Lotes.Add().Equals(0))
                    //{
                    //    Console.WriteLine($"Datos agregados a tabla salidas lotes exitosamente {item.ItemCode} - {item.TipDoc} - {item.NumDoc} - {item.NumLot}");
                    //}
                    //else
                    //{
                    //    Console.WriteLine($"Error al agregar datos a tabla salidas de lotes. {oCompany.GetLastErrorDescription()}. {item.ItemCode} - {item.TipDoc} - {item.NumDoc} - {item.NumLot}");
                    //    res = false;
                    //    break;
                    //}


                    ////Agregar a la lista de update los valores de cada linea
                    //Sb.Clear();
                    //query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.UpdateInfoConteo.sql")).ReadToEnd();
                    //query = Sb.AppendFormat(query, Math.Abs(item.CantLote), item.ItemCode, item.TipoDocumentoEntrada, item.NumeroDocumentoEntrada, item.NumLot, item.almacen).ToString();

                    //ListaUpdateEntrada.Add(query);
                    //Sb.Clear();





                }

                ////update de tabla de creacion lotes para descontar las cantidades por lote.
                //UpdateTablaEntradaLotes(ListaUpdateEntrada);

                ListaConteo.Clear();
                return res;

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en crear lotes de articulos con lote Conteo, {ex.Message}");
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(Lotes);

            }


        }

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
                        Console.WriteLine("Error al ejecutar recordset de update entrada lotes Conteo");
                    }
                    else
                    {
                        Console.WriteLine("Exito al actualizar entrada lotes Conteo");
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

        public void formClose()
        {
            try
            {
                ListaConteo.Clear();

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en formClose lotes Conteo de articulos con lote entrada, {ex.Message}");
            }
            finally
            {

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

                if (!Salida.Items.Item("BtnAgre").Enabled)
                {
                    return;
                }
                grid_Disponible = (Grid)Salida.Items.Item("GrdCrear").Specific;
                DT_Disponible = (DataTable)Salida.DataSources.DataTables.Item("DTCrea");
                grid_LineasDoc = (Grid)Salida.Items.Item("GrdLineas").Specific;
                grid_Seleccion = (Grid)Salida.Items.Item("GrdSel").Specific;
                DT_Seleccion = (DataTable)Salida.DataSources.DataTables.Item("DTSel");

                lineaseleccionada = ListaConteo.Count;


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
                            DateTime ofechaAdim = Convert.ToDateTime(fechaAdmin);

                            query = Sb.AppendFormat(query, DT_Disponible.GetValue("ItemCode", i), DT_Disponible.GetValue("Tipo doc", i), DT_Disponible.GetValue("Num Doc", i), DT_Disponible.GetValue("Numero lote", i), DT_Disponible.GetValue("Cantidad", i), ofechaVence.ToString("dd/MM/yyyy"), ofechaAdim.ToString("dd/MM/yyyy"), DT_Disponible.GetValue("Almacen", i), lineaseleccionada, i).ToString();
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
        /// <param name="pVal"></param>
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

                if (!Salida.Items.Item("BtnElim").Enabled)
                {
                    return;
                }

                grid_Disponible = (Grid)Salida.Items.Item("GrdCrear").Specific;
                DT_Disponible = (DataTable)Salida.DataSources.DataTables.Item("DTCrea");
                grid_Seleccion = (Grid)Salida.Items.Item("GrdSel").Specific;
                DT_Seleccion = (DataTable)Salida.DataSources.DataTables.Item("DTSel");

                for (int i = 0; i < grid_Seleccion.Rows.Count; i++)
                {
                    if (grid_Seleccion.Rows.IsSelected(i))
                    {

                        if (ListaConteo.Any())
                        {
                            ListaConteo.RemoveAll(x =>
                                                    x.ItemCode == DT_Seleccion.GetValue("ItemCode", i).ToString() &&
                                                    x.NumLot == DT_Seleccion.GetValue("Numero lote", i).ToString() &&
                                                    x.almacen == DT_Seleccion.GetValue("Almacen", i).ToString() &&
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

        internal bool CrearLotesContablizacion(BusinessObjectInfo businessObjectInfo, string Tipo)
        {
            bool res = true;
            SAPbobsCOM.UserTable Lotes = null;
            SAPbobsCOM.Recordset oRs = null;
            SAPbouiCOM.Form contabilizacion = null;
            SAPbouiCOM.Matrix LineasContabilizacion = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.EditText oCantidad = null;
            SAPbouiCOM.EditText oAlmacen = null;
            SAPbouiCOM.EditText oDocNum = null;
            SAPbouiCOM.EditText oDocNumConteo = null;
            SAPbouiCOM.EditText RowNumber = null;

            List<string> ListaUpdateEntrada = new List<string>();
            List<string> ListaDeleteConteo = new List<string>();
            var Sb = new System.Text.StringBuilder();
            string query = string.Empty, queryDelete = string.Empty;
            List<string> ItemsConteo = new List<string>();

            double cantidad = 0;

            //toca modificarlo para que consulte la tabla temporal
            try
            {
                contabilizacion = oApp.Forms.Item(businessObjectInfo.FormUID);

                if (!contabilizacion.Mode.Equals(BoFormMode.fm_ADD_MODE)) return true;



                Sb.Clear();
                query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.InfoConteo.sql")).ReadToEnd();

                LineasContabilizacion = (SAPbouiCOM.Matrix)contabilizacion.Items.Item("1470000022").Specific;

                for (int i = 1; i <= LineasContabilizacion.RowCount; i++)
                {
                    oEditText = (SAPbouiCOM.EditText)LineasContabilizacion.Columns.Item("1470000001").Cells.Item(i).Specific;



                    if (!string.IsNullOrEmpty(oEditText.String))
                    {


                        oAlmacen = (SAPbouiCOM.EditText)LineasContabilizacion.Columns.Item("1470000005").Cells.Item(i).Specific;

                        oCantidad = (SAPbouiCOM.EditText)LineasContabilizacion.Columns.Item("1470000009").Cells.Item(i).Specific;

                        oDocNum = (SAPbouiCOM.EditText)contabilizacion.Items.Item("1470000011").Specific;

                        oDocNumConteo = (SAPbouiCOM.EditText)LineasContabilizacion.Columns.Item("1470000043").Cells.Item(i).Specific;

                        RowNumber = (SAPbouiCOM.EditText)LineasContabilizacion.Columns.Item("1470000045").Cells.Item(i).Specific;

                        cantidad = Ini.ConvertToSystemFormat(oCantidad.String);
                        Sb.Clear();
                        ItemsConteo.Add(Sb.AppendFormat(query, oEditText.String, oDocNum.String, oAlmacen.String, Tipo, RowNumber.Value, oDocNumConteo.String, cantidad).ToString());
                        Sb.Clear();
                        System.Console.WriteLine(oEditText.String);

                    }
                }

                query = string.Join(" UNION ALL ", ItemsConteo);
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRs.DoQuery(query);

                if (oRs == null)
                {
                    oApp.StatusBar.SetText("Error en llamado de conteo temporal, para cargar datos en tabla de salidad o entradas", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else
                {
                    oRs.MoveFirst();
                    for (int i = 1; i <= oRs.RecordCount; i++)
                    {

                        if (oRs.Fields.Item("U_TipoCont").Value.ToString().Equals("E"))
                        {
                            Lotes = oCompany.UserTables.Item("CLC_LTS");

                            Lotes.Name = oRs.Fields.Item("Name").Value.ToString();
                            Lotes.UserFields.Fields.Item("U_TipDoc").Value = oRs.Fields.Item("TipDoc").Value.ToString();
                            Lotes.UserFields.Fields.Item("U_NumDoc").Value = oRs.Fields.Item("NumDoc").Value.ToString();
                            Lotes.UserFields.Fields.Item("U_NumLot").Value = oRs.Fields.Item("U_NumLot").Value.ToString();
                            Lotes.UserFields.Fields.Item("U_CantLote").Value = oRs.Fields.Item("U_CantLote").Value;
                            Lotes.UserFields.Fields.Item("U_FchVenLt").Value = oRs.Fields.Item("U_FchVenLt").Value.ToString();
                            Lotes.UserFields.Fields.Item("U_FchAdmLt").Value = oRs.Fields.Item("U_FchAdmLt").Value.ToString();
                            Lotes.UserFields.Fields.Item("U_almacen").Value = oRs.Fields.Item("U_almacen").Value.ToString();

                            if (Lotes.Add().Equals(0))
                            {
                                Console.WriteLine($"Datos agregados a tabla entradas lotes exitosamente {oRs.Fields.Item("Name").Value.ToString()} - {oRs.Fields.Item("TipDoc").Value.ToString()} - {oRs.Fields.Item("NumDoc").Value.ToString()} - {oRs.Fields.Item("U_NumLot").Value.ToString()}");
                                Sb.Clear();
                                queryDelete = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.DeleteConteo.sql")).ReadToEnd();
                                queryDelete = Sb.AppendFormat(queryDelete, oRs.Fields.Item("Name").Value.ToString(), oRs.Fields.Item("U_NmDocEnt").Value.ToString(), oRs.Fields.Item("U_almacen").Value.ToString(), oRs.Fields.Item("U_RowNum").Value.ToString()).ToString();
                                ListaDeleteConteo.Add(queryDelete);
                            }
                            else
                            {
                                Console.WriteLine($"Error al agregar datos a tabla entrada de lotes. {oCompany.GetLastErrorDescription()}. {oRs.Fields.Item("Name").Value.ToString()} - {oRs.Fields.Item("TipDoc").Value.ToString()} - {oRs.Fields.Item("NumDoc").Value.ToString()} - {oRs.Fields.Item("U_NumLot").Value.ToString()}");
                                res = false;
                                break;
                            }
                        }
                        else
                        {
                            Lotes = oCompany.UserTables.Item("CLC_LTSSA");

                            Lotes.Name = oRs.Fields.Item("Name").Value.ToString();
                            Lotes.UserFields.Fields.Item("U_TipDoc").Value = oRs.Fields.Item("TipDoc").Value.ToString();
                            Lotes.UserFields.Fields.Item("U_NumDoc").Value = oRs.Fields.Item("NumDoc").Value.ToString();
                            Lotes.UserFields.Fields.Item("U_NumLot").Value = oRs.Fields.Item("U_NumLot").Value.ToString();
                            Lotes.UserFields.Fields.Item("U_Cantidad").Value = oRs.Fields.Item("U_CantLote").Value;
                            Lotes.UserFields.Fields.Item("U_FchSal").Value = DateTime.Now;
                            Lotes.UserFields.Fields.Item("U_almacen").Value = oRs.Fields.Item("U_almacen").Value.ToString();

                            if (Lotes.Add().Equals(0))
                            {
                                Console.WriteLine($"Datos agregados a tabla salidas lotes exitosamente {oRs.Fields.Item("Name").Value.ToString()} - {oRs.Fields.Item("TipDoc").Value.ToString()} - {oRs.Fields.Item("NumDoc").Value.ToString()} - {oRs.Fields.Item("U_NumLot").Value.ToString()}");
                                Sb.Clear();
                                queryDelete = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.DeleteConteo.sql")).ReadToEnd();
                                queryDelete = Sb.AppendFormat(queryDelete, oRs.Fields.Item("Name").Value.ToString(), oRs.Fields.Item("NumDoc2").Value.ToString(), oRs.Fields.Item("U_almacen").Value.ToString(), oRs.Fields.Item("U_RowNum").Value.ToString()).ToString();
                                ListaDeleteConteo.Add(queryDelete);

                            }
                            else
                            {
                                Console.WriteLine($"Error al agregar datos a tabla salidas de lotes. {oCompany.GetLastErrorDescription()}. {oRs.Fields.Item("Name").Value.ToString()} - {oRs.Fields.Item("TipDoc").Value.ToString()} - {oRs.Fields.Item("NumDoc").Value.ToString()} - {oRs.Fields.Item("U_NumLot").Value.ToString()}");
                                res = false;
                                break;
                            }


                            //Agregar a la lista de update los valores de cada linea
                            Sb.Clear();
                            query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.UpdateInfoConteo.sql")).ReadToEnd();
                            query = Sb.AppendFormat(query, oRs.Fields.Item("U_CantLote").Value, oRs.Fields.Item("Name").Value.ToString(), oRs.Fields.Item("U_TpDocEnt").Value.ToString(), oRs.Fields.Item("U_NmDocEnt").Value.ToString(), oRs.Fields.Item("U_NumLot").Value.ToString(), oRs.Fields.Item("U_almacen").Value.ToString()).ToString();

                            ListaUpdateEntrada.Add(query);
                            Sb.Clear();


                        }
                        oRs.MoveNext();
                    }

                }



                //update de tabla de creacion lotes para descontar las cantidades por lote.
                UpdateTablaEntradaLotes(ListaUpdateEntrada);
                DeleteConteos(ListaDeleteConteo);

                return res;

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en crear lotes de articulos con lote Conteo, {ex.Message}");
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(Lotes);
                Ini.ReleaseCOMObjetc(oRs);
                Ini.ReleaseCOMObjetc(LineasContabilizacion);
                Ini.ReleaseCOMObjetc(oEditText);
                Ini.ReleaseCOMObjetc(oAlmacen);
                Ini.ReleaseCOMObjetc(oDocNum);
                Ini.ReleaseCOMObjetc(oDocNumConteo);
                Ini.ReleaseCOMObjetc(RowNumber);

            }
        }

        private void DeleteConteos(List<string> listaDeleteConteo)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                foreach (var lista in listaDeleteConteo)
                {
                    oRs.DoQuery(lista);
                    if (oRs == null)
                    {
                        Console.WriteLine("Error al ejecutar recordset de delete Conteo");
                    }
                    else
                    {
                        Console.WriteLine("Exito al ejecutar recordset de delete Conteo");
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

        internal void AddLine(ItemEvent pVal)
        {
            SAPbouiCOM.Form oForm = null;
            DataTable DT_Crear = null;
            Grid grid_Crear = null;


            try
            {
                oForm = oApp.Forms.Item(pVal.FormUID);
                if (oForm.Items.Item("BtAddln").Enabled.Equals(false)) return;

                oForm.Freeze(true);
                DT_Crear = (DataTable)oForm.DataSources.DataTables.Item("DTCrea");
                grid_Crear = (Grid)oForm.Items.Item("GrdCrear").Specific;



                DT_Crear.Rows.Add();
                DT_Crear.SetValue("ItemCode", grid_Crear.Rows.Count - 1, DT_Crear.GetValue("ItemCode", grid_Crear.Rows.Count - 2));
                DT_Crear.SetValue("Cantidad", grid_Crear.Rows.Count - 1, DT_Crear.GetValue("Cantidad", grid_Crear.Rows.Count - 2));
                DT_Crear.SetValue("Fecha vence", grid_Crear.Rows.Count - 1, DT_Crear.GetValue("Fecha vence", grid_Crear.Rows.Count - 2));
                DT_Crear.SetValue("Fecha Admin", grid_Crear.Rows.Count - 1, DT_Crear.GetValue("Fecha Admin", grid_Crear.Rows.Count - 2));
                DT_Crear.SetValue("Tipo doc", grid_Crear.Rows.Count - 1, DT_Crear.GetValue("Tipo doc", grid_Crear.Rows.Count - 2));
                DT_Crear.SetValue("Num Doc", grid_Crear.Rows.Count - 1, DT_Crear.GetValue("Num Doc", grid_Crear.Rows.Count - 2));
                DT_Crear.SetValue("Almacen", grid_Crear.Rows.Count - 1, DT_Crear.GetValue("Almacen", grid_Crear.Rows.Count - 2));
                DT_Crear.SetValue("NumLinea", grid_Crear.Rows.Count - 1, DT_Crear.GetValue("NumLinea", grid_Crear.Rows.Count - 2));

            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                Console.WriteLine($"Error en agregar línea {ex.Message}");

            }
            finally
            {
                oForm.Freeze(false);
                Ini.ReleaseCOMObjetc(oForm);
                Ini.ReleaseCOMObjetc(DT_Crear);
                Ini.ReleaseCOMObjetc(grid_Crear);
            }
        }

        internal void DeleteRow(ItemEvent pVal)
        {
            DataTable DTCrear = null;
            SAPbouiCOM.Form oForm = null;



            try
            {
                oForm = oApp.Forms.Item(pVal.FormUID);
                DTCrear = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DTCrea");

                for (int i = 0; i < DTCrear.Rows.Count; i++)
                {
                    if (DTCrear.GetValue("Sel", i).Equals("Y"))
                    {
                        DTCrear.Rows.Remove(i);
                        if (ListaConteo.Any())
                        {
                            ListaConteo.RemoveAt((int)DTCrear.GetValue("Index", i));

                        }

                    }
                }

                if (DTCrear.IsEmpty)
                {

                }

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en  DeleteRow de articulos con lote entrada, {ex.Message}");
            }
            finally
            {
                Ini.ReleaseCOMObjetc(DTCrear);
                Ini.ReleaseCOMObjetc(oForm);


            }

        }
    }
}